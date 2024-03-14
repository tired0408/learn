import time
import random
import torch
from torch import nn
from torch import optim
import gym
import cv2
import numpy as np
import matplotlib.pyplot as plt
from collections import deque, namedtuple
from tqdm import tqdm

device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
Transition = namedtuple('Transition', ('state', 'action', 'reward', 'next_state', 'done'))


def show_image(image):
    cv2.imshow("img", image)
    return cv2.waitKey(0) & 0xFF


class ReplayMemory:

    def __init__(self, memory_size):
        self.memory = deque([], maxlen=memory_size)

    def sample(self, batch_size):
        batch_data = random.sample(self.memory, batch_size)
        state, action, reward, next_state, done = zip(*batch_data)
        return state, action, reward, next_state, done

    def push(self, *args):
        self.memory.append(Transition(*args))

    def __len__(self):
        return len(self.memory)


class Qnet(nn.Module):

    def __init__(self, n_observations, n_actions):
        super(Qnet, self).__init__()
        self.model = nn.Sequential(
            nn.Linear(n_observations, 128),
            nn.ReLU(),
            nn.Linear(128, n_actions)
        )

    def forward(self, state):
        return self.model(state)


class Agent(object):

    def __init__(self, observation_dim, action_dim, gamma, lr, epsilon, target_update):
        self.action_dim = action_dim
        self.q_net = Qnet(observation_dim, action_dim).to(device)
        self.target_q_net = Qnet(observation_dim, action_dim).to(device)
        self.gamma = gamma
        self.lr = lr
        self.epsilon = epsilon
        self.target_update = target_update
        self.count = 0

        self.optimizer = optim.Adam(params=self.q_net.parameters(), lr=lr)
        self.loss = nn.MSELoss()

    def take_action(self, state):
        if np.random.uniform(0, 1) < 1 - self.epsilon:
            state = torch.tensor(state, dtype=torch.float).to(device)
            action = torch.argmax(self.q_net(state)).item()
        else:
            action = np.random.choice(self.action_dim)
        return action

    def update(self, transition_dict):

        states = transition_dict.state
        actions = np.expand_dims(transition_dict.action, axis=-1)  # 扩充维度
        rewards = np.expand_dims(transition_dict.reward, axis=-1)  # 扩充维度
        next_states = transition_dict.next_state
        dones = np.expand_dims(transition_dict.done, axis=-1)  # 扩充维度

        states = torch.tensor(states, dtype=torch.float).to(device)
        actions = torch.tensor(actions, dtype=torch.int64).to(device)
        rewards = torch.tensor(rewards, dtype=torch.float).to(device)
        next_states = torch.tensor(next_states, dtype=torch.float).to(device)
        dones = torch.tensor(dones, dtype=torch.float).to(device)

        # update q_values
        # gather(1, acitons)意思是dim=1按行号索引， index=actions
        # actions=[[1, 2], [0, 1]] 意思是索引出[[第一行第2个元素， 第1行第3个元素],[第2行第1个元素， 第2行第2个元素]]
        # 相反，如果是这样
        # gather(0, acitons)意思是dim=0按列号索引， index=actions
        # actions=[[1, 2], [0, 1]] 意思是索引出[[第一列第2个元素， 第2列第3个元素],[第1列第1个元素， 第2列第2个元素]]
        # states.shape(64, 4) actions.shape(64, 1), 每一行是一个样本，所以这里用dim=1很合适
        predict_q_values = self.q_net(states).gather(1, actions)
        with torch.no_grad():
            # max(1) 即 max(dim=1)在行向找最大值，这样的话shape(64, ), 所以再加一个view(-1, 1)扩增至(64, 1)
            max_next_q_values = self.target_q_net(next_states).max(1)[0].view(-1, 1)
            q_targets = rewards + self.gamma * max_next_q_values * (1 - dones)
        loss_value = self.loss(predict_q_values, q_targets)

        self.optimizer.zero_grad()
        loss_value.backward()
        self.optimizer.step()

        if self.count % self.target_update == 0:
            # copy model parameters
            self.target_q_net.load_state_dict(self.q_net.state_dict())

        self.count += 1


def run_episode(env, agent, repalymemory, batch_size):
    state = env.reset()[0]
    reward_total = 0
    while True:
        action = agent.take_action(state)
        next_state, reward, done, _, _ = env.step(action)
        repalymemory.push(state, action, reward, next_state, done)
        reward_total += reward
        if len(repalymemory) > batch_size:
            state_batch, action_batch, reward_batch, next_state_batch, done_batch = repalymemory.sample(batch_size)
            T_data = Transition(state_batch, action_batch, reward_batch, next_state_batch, done_batch)
            agent.update(T_data)
        state = next_state
        if done:
            break
    return reward_total


def episode_evaluate(env, agent, render):
    reward_list = []
    for i in range(5):
        state = env.reset()[0]
        reward_episode = 0
        while True:
            action = agent.take_action(state)
            next_state, reward, done, _, _ = env.step(action)
            reward_episode += reward
            state = next_state
            if done:
                break
            if render:
                env.render()
        reward_list.append(reward_episode)
    return np.mean(reward_list).item()


def test(env, agent):
    state = env.reset()[0]
    reward_episode = 0
    while True:
        action = agent.take_action(state)
        next_state, reward, done, _, _ = env.step(action)
        reward_episode += reward
        state = next_state
        print(f"action:{action}, reward:{reward}")
        img = env.render()
        show_image(img)
        if done:
            break


if __name__ == "__main__":
    env = gym.make("CartPole-v0", render_mode="rgb_array")
    env_name = "CartPole-v0"
    observation_n, action_n = env.observation_space.shape[0], env.action_space.n

    agent = Agent(observation_n, action_n, gamma=0.98, lr=2e-3, epsilon=0.01, target_update=10)

    replaymemory = ReplayMemory(memory_size=10000)
    batch_size = 64

    num_episodes = 200
    reward_list = []
    # 显示10个进度条
    for i in range(10):
        with tqdm(total=int(num_episodes / 10), desc="Iteration %d" % i) as pbar:
            for episode in range(int(num_episodes / 10)):

                reward_episode = run_episode(env, agent, replaymemory, batch_size)
                reward_list.append(reward_episode)

                if (episode + 1) % 10 == 0:

                    test_reward = episode_evaluate(env, agent, False)
                    # print("Episode %d, total reward: %.3f" % (episode, test_reward))
                    pbar.set_postfix({
                        'episode': '%d' % (num_episodes / 10 * i + episode + 1),
                        'return': '%.3f' % (test_reward)
                    })
                pbar.update(1)  # 更新进度条

    test(env, agent)     # 最后用动画观看一下效果
    episodes_list = list(range(len(reward_list)))
    plt.plot(episodes_list, reward_list)
    plt.xlabel('Episodes')
    plt.ylabel('Returns')
    plt.title('Double DQN on {}'.format(env_name))
    plt.show()
