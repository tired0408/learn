"""
QLearning算法解决悬崖寻路问题
"""
import gym
import math
import torch
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from collections import defaultdict


class Config:

    def __init__(self):
        self.env_name = "CliffWalking-v0"
        self.algo_name = "Q-Learning"
        self.train_eps = 400
        self.test_eps = 20
        self.max_steps = 200
        self.epsilon_start = 0.95
        self.epsilon_end = 0.01
        self.epsilon_decay = 300
        self.gamma = 0.9
        self.lr = 0.1
        self.seed = 1
        self.device = torch.device("cuda") if torch.cuda.is_available() else torch.device("cpu")


class QLearning:

    def __init__(self, n_actions, cfg: Config):
        self.n_actions = n_actions
        self.lr = cfg.lr
        self.gamma = cfg.gamma
        self.epsilon = cfg.epsilon_start
        self.sample_count = 0
        self.epsilon_start = cfg.epsilon_start
        self.epsilon_end = cfg.epsilon_end
        self.epsilon_decay = cfg.epsilon_decay
        self.Q_table = defaultdict(lambda: np.zeros(n_actions))

    def sample_action(self, state):
        """采样动作"""
        self.sample_count += 1
        self.epsilon = self.epsilon_end + (self.epsilon_start - self.epsilon_end) * \
            math.exp(-1 * self.sample_count / self.epsilon_decay)
        if np.random.uniform(0, 1) > self.epsilon:
            action = np.argmax(self.Q_table[str(state)])
        else:
            action = np.random.choice(self.n_actions)
        return action, self.epsilon

    def predict_action(self, state):
        """预测或选择动作，测试时用"""
        action = np.argmax(self.Q_table[str(state)])
        return action

    def update(self, state, action, reward, next_state, terminated):
        Q_predict = self.Q_table[str(state)][action]
        if terminated:
            Q_target = reward
        else:
            Q_target = reward + self.gamma * np.max(self.Q_table[str(next_state)])
        self.Q_table[str(state)][action] += self.lr * (Q_target - Q_predict)


def env_agent_config(cfg: Config):
    """创建环境和智能体"""
    env = gym.make(cfg.env_name)
    n_actions = env.action_space.n
    agent = QLearning(n_actions, cfg)
    return env, agent


def train(cfg: Config, env: gym.Env, agent: QLearning):
    print("开始训练")
    print(f"环境:{cfg.env_name}, 算法:{cfg.algo_name}, 设备:{cfg.device}")
    rewards = []
    for i_ep in range(cfg.train_eps):
        ep_reward = 0
        state = env.reset(seed=cfg.seed)
        while True:
            # env.render()
            action, epsilon = agent.sample_action(state)
            next_state, reward, terminated, _ = env.step(action)
            agent.update(state, action, reward, next_state, terminated)
            # print(state, action, reward, next_state, terminated, epsilon)
            state = next_state
            ep_reward += reward
            if terminated:
                break
        rewards.append(ep_reward)
        if (i_ep + 1) % 10 == 0:
            print(f"Episode:{i_ep + 1}/{cfg.train_eps}, Reward:{ep_reward:.1f}, Epsilon:{agent.epsilon:.3f}")
    print('完成训练')
    return {"rewards": rewards}


def test(cfg: Config, env: gym.Env, agent: QLearning):
    print("开始测试")
    print(f"环境:{cfg.env_name}, 算法:{cfg.algo_name}, 设备:{cfg.device}")
    rewards = []
    for i_ep in range(cfg.test_eps):
        ep_reward = 0
        state = env.reset(seed=cfg.seed)
        while True:
            env.render()
            action = agent.predict_action(state)
            next_state, reward, terminated, _ = env.step(action)
            state = next_state
            ep_reward += reward
            if terminated:
                break
        rewards.append(ep_reward)
        print(f"Episode:{i_ep + 1}/{cfg.test_eps}, Reward:{ep_reward:.1f}")
    print('完成测试')
    return {"rewards": rewards}


def smooth(data, weight=0.9):
    '''用于平滑曲线,类似于Tensorboard中的smooth

    Args:
        data (List):输入数据
        weight (Float): 平滑权重,处于0-1之间,数值越高说明越平滑一般取0.9

    Returns:
        smoothed (List): 平滑后的数据
    '''
    last = data[0]  # First value in the plot (first timestep)
    smoothed = list()
    for point in data:
        smoothed_val = last * weight + (1 - weight) * point  # 计算平滑值
        smoothed.append(smoothed_val)
        last = smoothed_val
    return smoothed


def plot_rewards(rewards, title):
    sns.set()
    plt.figure()  # 创建一个图形实例，方便同时多画几个图
    plt.title(title)
    plt.xlabel('epsiodes')
    plt.plot(rewards, label='rewards')
    plt.plot(smooth(rewards), label='smoothed')
    plt.legend()
    plt.show()


if __name__ == "__main__":
    cfg = Config()
    env, agent = env_agent_config(cfg)
    res_dic = train(cfg, env, agent)
    plot_rewards(res_dic['rewards'], title=f"training curve on {cfg.device} of {cfg.algo_name} for {cfg.env_name}")
    # 测试
    res_dic = test(cfg, env, agent)
    plot_rewards(res_dic['rewards'], title=f"testing curve on {cfg.device} of {cfg.algo_name} for {cfg.env_name}")
