"""
Double-DQN的算法实践:解决CartPole的问题
"""
import random
from collections import deque
from torch import nn
from torch.nn import functional as F


class MLP(nn.module):
    """定义多层感知器神经网络"""

    def __init__(self, n_states, n_actions, hidden_dim=128):
        super().__init__()
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
