"""
DQN的算法实践:解决CartPole的问题
"""
import os
import gym
import random
import torch
import math
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from torch import optim
from collections import deque
from torch import nn
from torch.nn import functional as F


class Config:

    def __init__(self):
        self.algo_name: str = "DQN"
        self.env_name: str = "CartPole-v0"
        self.train_eps: int = 200
        self.test_eps: int = 20
        self.ep_max_steps: int = 100000
        self.gamma: float = 0.95
        self.epsilon_start: float = 0.95
        self.epsilon_end: float = 0.01
        self.epsilon_decay: int = 500
        self.lr: float = 0.0001
        self.memory_capacity: int = 100000
        self.batch_size: int = 64
        self.target_update: int = 4
        self.hidden_dim: int = 256
        self.device: str = "cpu"
        self.seed: int = 10
        self.n_states: int = 0
        self.n_actions: int = 0

    def update(self, **kwargs):
        for k, v in kwargs.items():
            setattr(self, k, v)

    def print_params(self):
        print("超参数")
        print("=" * 80)
        tplt = "{:^20}\t{:^20}\t{:^20}"
        print(tplt.format("Name", "Value", "Type"))
        for k, v in vars(self).items():
            print(tplt.format(k, v, str(type(v))))
        print("=" * 80)


class MLP(nn.Module):
    """定义多层感知器神经网络"""

    def __call__(self, *args, **kwds) -> torch.Tensor:
        return super().__call__(*args, **kwds)

    def __init__(self, n_states, n_actions, hidden_dim=128):
        super(MLP, self).__init__()
        self.fc1 = nn.Linear(n_states, hidden_dim)
        self.fc2 = nn.Linear(hidden_dim, hidden_dim)
        self.fc3 = nn.Linear(hidden_dim, n_actions)

    def forward(self, x):
        x = F.relu(self.fc1(x))
        x = F.relu(self.fc2(x))
        return self.fc3(x)


class ReplayBuffer:

    def __init__(self, capacity: int):
        self.capacity = capacity
        self.buffer = deque(maxlen=self.capacity)

    def push(self, transitions):
        """存储transition到经验回放中"""
        self.buffer.append(transitions)

    def sample(self, batch_size: int, sequential: bool = False):
        """从经验回放中进行采样"""
        if batch_size > len(self.buffer):
            batch_size = len(self.buffer)
        if sequential:
            rand = random.randint(0, len(self.buffer) - batch_size)
            batch = [self.buffer[i] for i in range(rand, rand + batch_size)]
            return zip(*batch)
        else:
            batch = random.sample(self.buffer, batch_size)
            return zip(*batch)

    def clear(self):
        self.buffer.clear()

    def __len__(self):
        return len(self.buffer)


class DNN:

    def __init__(self, model: MLP, memory: ReplayBuffer, cfg: Config):
        self.n_actions = cfg.n_actions
        self.device = torch.device(cfg.device)
        self.gamma = cfg.gamma
        # e-greedy策略相关参数
        self.epsilon = cfg.epsilon_start
        self.sample_count = 0
        self.epsilon_start = cfg.epsilon_start
        self.epsilon_end = cfg.epsilon_end
        self.epsilon_decay = cfg.epsilon_decay
        # 神经网络参数
        self.batch_size = cfg.batch_size
        self.policy_net = model.to(self.device)
        self.target_net = model.to(self.device)
        self.target_net.load_state_dict(self.policy_net.state_dict())
        self.optimizer = optim.Adam(self.policy_net.parameters(), lr=cfg.lr)
        self.memory = memory

    def sample_action(self, state):
        """采样动作"""
        self.sample_count += 1
        self.epsilon = self.epsilon_end + (self.epsilon_start - self.epsilon_end) * \
            math.exp(-1. * self.sample_count / self.epsilon_decay)
        if random.random() > self.epsilon:
            with torch.no_grad():
                state = torch.tensor(state, device=self.device, dtype=torch.float32).unsqueeze(0)
                q_values = self.policy_net(state)
                action = q_values.max(1)[1].item()
        else:
            action = random.randrange(self.n_actions)
        return action

    @ torch.no_grad()
    def predict_action(self, state):
        """预测动作"""
        state = torch.tensor(state, device=self.device, dtype=torch.float32).unsqueeze(0)
        q_values = self.policy_net(state)
        action = q_values.max(1)[1].item()
        return action

    def update(self):
        """更新参数"""
        if len(self.memory) < self.batch_size:
            return
        state_batch, action_batch, reward_batch, next_state_batch, done_batch = self.memory.sample(self.batch_size)
        state_batch = torch.tensor(np.array(state_batch), device=self.device, dtype=torch.float)
        action_batch = torch.tensor(action_batch, device=self.device).unsqueeze(1)
        reward_batch = torch.tensor(reward_batch, device=self.device, dtype=torch.float)
        next_state_batch = torch.tensor(np.array(next_state_batch), device=self.device, dtype=torch.float)
        done_batch = torch.tensor(np.float32(done_batch), device=self.device)
        q_values = self.policy_net(state_batch).gather(dim=1, index=action_batch)
        next_q_values = self.target_net(next_state_batch).max(1)[0].detach()
        expected_q_values = reward_batch + self.gamma * next_q_values * (1 - done_batch)
        loss: torch.Tensor = nn.MSELoss()(q_values, expected_q_values.unsqueeze(1))
        self.optimizer.zero_grad()
        loss.backward()
        for param in self.policy_net.parameters():
            param.grad.data.clamp_(-1, 1)
        self.optimizer.step()


def train(cfg: Config, env: gym.Env, agent: DNN):
    """训练"""
    print("开始训练")
    rewards = []
    steps = []
    for i_ep in range(cfg.train_eps):
        ep_reward = 0
        ep_step = 0
        state = env.reset()
        for _ in range(cfg.ep_max_steps):
            ep_step += 1
            action = agent.sample_action(state)
            next_state, reward, done, _ = env.step(action)
            agent.memory.push((state, action, reward, next_state, done))
            state = next_state
            agent.update()
            ep_reward += reward
            if done:
                break
        if (i_ep + 1) % cfg.target_update == 0:
            agent.target_net.load_state_dict(agent.policy_net.state_dict())
        steps.append(ep_step)
        rewards.append(ep_reward)
        if (i_ep + 1) % 10 == 0:
            print(f"回合:{i_ep + 1:^3}/{cfg.train_eps}, 奖励:{f'{ep_reward:.2f},':^8} Epislon:{agent.epsilon:.3f}")
    print("完成训练")
    env.close()
    return {"rewards": rewards}


def test(cfg: Config, env: gym.Env, agent: DNN):
    """测试"""
    print("开始测试")
    rewards = []
    steps = []
    for i_ep in range(cfg.test_eps):
        ep_reward = 0
        ep_step = 0
        state = env.reset()
        for _ in range(cfg.ep_max_steps):
            env.render()
            ep_step += 1
            action = agent.predict_action(state)
            next_state, reward, done, _ = env.step(action)
            # print(f"执行动作:{'向左' if action == 0 else '向右'},下一个状态:{next_state}, 奖励:{reward}, 是否结束:{done}")
            state = next_state
            ep_reward += reward
            if done:
                break
        steps.append(ep_step)
        rewards.append(ep_reward)
        print(f"回合:{i_ep + 1:^3}/{cfg.test_eps}, 奖励:{ep_reward:.2f}")
    print("完成测试")
    env.close()
    return {"rewards": rewards}


def all_seed(env: gym.Env, seed=1):
    """万能的seed函数"""
    env.seed(seed)
    np.random.seed(seed)
    random.seed(seed)
    torch.manual_seed(seed)
    torch.cuda.manual_seed(seed)
    os.environ['PYTHONHASHSEED'] = str(seed)
    torch.backends.cudnn.deterministic = True
    torch.backends.cudnn.benchmark = False
    torch.backends.cudnn.enabled = False


def env_agent_config(cfg: Config):
    """创建环境和智能体"""
    env = gym.make(cfg.env_name)
    if cfg.seed != 0:
        all_seed(env, seed=cfg.seed)
    n_states = env.observation_space.shape[0]
    n_actions = env.action_space.n
    print(f"状态数:{n_states}, 动作数:{n_actions}")
    cfg.update(n_states=n_states, n_actions=n_actions)
    model = MLP(n_states, n_actions, cfg.hidden_dim)
    memory = ReplayBuffer(cfg.memory_capacity)
    agent = DNN(model, memory, cfg)
    return env, agent


def smooth(data, weight=0.9):
    """用于平滑曲线"""
    last = data[0]
    smoothed = []
    for point in data:
        smoothed_val = last * weight + (1 - weight) * point
        smoothed.append(smoothed_val)
        last = smoothed_val
    return smoothed


def plot_rewards(rewards, tag="train"):
    """绘制奖励曲线"""
    sns.set()
    plt.figure()
    plt.title(f"{tag}ing curve")
    plt.xlabel("epsiods")
    plt.plot(rewards, label="rewards")
    plt.plot(smooth(rewards), label="smoothed")
    plt.legend()
    plt.show()


if __name__ == "__main__":
    cfg = Config()
    cfg.print_params()
    env, agent = env_agent_config(cfg)
    res_dic = train(cfg, env, agent)
    plot_rewards(res_dic['rewards'], tag="train")
    res_dic = test(cfg, env, agent)
    plot_rewards(res_dic['rewards'], tag="test")
