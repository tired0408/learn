import numpy as np
import tensorflow as tf
import os
import pandas as pd
import numpy as np
import math
import time
import pickle
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2' #只显示warning和error
checkpoint_dir = ''
def load_datasets():
    path = "E:/tencent_data/"
    save_path = path + "deal_data/"
    df = pd.read_csv(save_path + "train_set_finall.csv", sep="\t", low_memory=False)
    df = df.reindex(np.random.permutation(df.index))
    train = df.iloc[20000:]
    test = df.iloc[0:20000]
    train_set_x_orig = np.array(train.drop(["label"], axis=1))
    train_set_y_orig = np.array(train[["label"]])
    test_set_x_orig = np.array(test.drop(["label"], axis=1))
    test_set_y_orig = np.array(test[["label"]])

    train_set_y_orig = train_set_y_orig.reshape((1, train_set_y_orig.shape[0]))
    test_set_y_orig = test_set_y_orig.reshape((1, test_set_y_orig.shape[0]))

    return train_set_x_orig, train_set_y_orig, test_set_x_orig, test_set_y_orig


def random_mini_batches(X, Y, mini_batch_size=64, seed=0):
    """
    Creates a list of random minibatches from (X, Y)
    Arguments:
    X -- input data, of shape (input size, number of examples)
    Y -- true "label" vector (containing 0 if cat, 1 if non-cat), of shape (1, number of examples)
    mini_batch_size - size of the mini-batches, integer
    seed -- this is only for the purpose of grading, so that you're "random minibatches are the same as ours.
    Returns:
    mini_batches -- list of synchronous (mini_batch_X, mini_batch_Y)
    """

    m = X.shape[1]  # number of training examples
    m = int(m)
    mini_batches = []
    np.random.seed(seed)

    # Step 1: Shuffle (X, Y)
    permutation = list(np.random.permutation(m))
    shuffled_X = X[:, permutation]
    shuffled_Y = Y[:, permutation].reshape((Y.shape[0], m))

    # Step 2: Partition (shuffled_X, shuffled_Y). Minus the end case.
    num_complete_minibatches = math.floor(
        m / mini_batch_size)  # number of mini batches of size mini_batch_size in your partitionning
    for k in range(0, num_complete_minibatches):
        mini_batch_X = shuffled_X[:, k * mini_batch_size: k * mini_batch_size + mini_batch_size]
        mini_batch_Y = shuffled_Y[:, k * mini_batch_size: k * mini_batch_size + mini_batch_size]
        mini_batch = (mini_batch_X, mini_batch_Y)
        mini_batches.append(mini_batch)

    # Handling the end case (last mini-batch < mini_batch_size)
    if m % mini_batch_size != 0:
        mini_batch_X = shuffled_X[:, num_complete_minibatches * mini_batch_size: m]
        mini_batch_Y = shuffled_Y[:, num_complete_minibatches * mini_batch_size: m]
        mini_batch = (mini_batch_X, mini_batch_Y)
        mini_batches.append(mini_batch)

    return mini_batches

def predict(X, parameters):
    W1 = tf.convert_to_tensor(parameters["W1"])
    b1 = tf.convert_to_tensor(parameters["b1"])
    W2 = tf.convert_to_tensor(parameters["W2"])
    b2 = tf.convert_to_tensor(parameters["b2"])
    W3 = tf.convert_to_tensor(parameters["W3"])
    b3 = tf.convert_to_tensor(parameters["b3"])

    params = {"W1": W1,
              "b1": b1,
              "W2": W2,
              "b2": b2,
              "W3": W3,
              "b3": b3}

    x = tf.placeholder("float", shape=[14, None])

    z3 = forward_propagation_for_predict(x, params)

    sess = tf.Session()
    prediction = sess.run(z3, feed_dict={x: X})

    return prediction


def forward_propagation_for_predict(X, parameters):
    """
    Implements the forward propagation for the model: LINEAR -> RELU -> LINEAR -> RELU -> LINEAR -> SOFTMAX
    Arguments:
    X -- input dataset placeholder, of shape (input size, number of examples)
    parameters -- python dictionary containing your parameters "W1", "b1", "W2", "b2", "W3", "b3"
                  the shapes are given in initialize_parameters
    Returns:
    Z3 -- the output of the last LINEAR unit
    """

    # Retrieve the parameters from the dictionary "parameters"
    W1 = parameters['W1']
    b1 = parameters['b1']
    W2 = parameters['W2']
    b2 = parameters['b2']
    W3 = parameters['W3']
    b3 = parameters['b3']
    # Numpy Equivalents:
    Z1 = tf.add(tf.matmul(W1, X), b1)  # Z1 = np.dot(W1, X) + b1
    A1 = tf.nn.relu(Z1)  # A1 = relu(Z1)
    Z2 = tf.add(tf.matmul(W2, A1), b2)  # Z2 = np.dot(W2, a1) + b2
    A2 = tf.nn.relu(Z2)  # A2 = relu(Z2)
    Z3 = tf.add(tf.matmul(W3, A2), b3)  # Z3 = np.dot(W3,Z2) + b3

    return Z3

def create_placeholders(n_x, n_y):
    X = tf.placeholder(tf.float32, shape=[n_x, None], name='X')
    Y = tf.placeholder(tf.float32, shape=[n_y, None], name='Y')
    return X,Y

def initialize_parameters(n_x,n_y):
    """
    初始化参数
    :param n_y:  标签的输出维度
    :param n_x: 训练样本的特征数
    :return:
    """
    W1 = tf.get_variable("W1", [25,14], initializer=tf.contrib.layers.xavier_initializer(seed=1))
    b1 = tf.get_variable("b1", [25,1], initializer=tf.zeros_initializer())
    W2 = tf.get_variable("W2", [12,25], initializer=tf.contrib.layers.xavier_initializer(seed=1))
    b2 = tf.get_variable("b2", [12,1], initializer=tf.zeros_initializer())
    W3 = tf.get_variable("W3", [1,12], initializer=tf.contrib.layers.xavier_initializer(seed=1))
    b3 = tf.get_variable("b3", [1,1], initializer=tf.zeros_initializer())

    parameters = {"W1":W1, "b1":b1, "W2":W2, "b2":b2, "W3":W3, "b3":b3}
    return parameters

def forward_propagation(X, parameters):
    W1 = parameters["W1"]
    b1 = parameters["b1"]
    W2 = parameters["W2"]
    b2 = parameters["b2"]
    W3 = parameters["W3"]
    b3 = parameters["b3"]

    Z1 = tf.matmul(W1,X) + b1
    A1 = tf.nn.relu(Z1)
    Z2 = tf.matmul(W2,A1) + b2
    A2 = tf.nn.relu(Z2)
    Z3 = tf.matmul(W3,A2) + b3
    return Z3

def compute_cost(Z3, Y):
    # 损失函数，这里采用的是最小二乘法的损失函数，即计算模型输出值与真实值之间的误差的最小二乘法
    loss = tf.reduce_mean(tf.reduce_sum(tf.square((Z3 - Y)), reduction_indices=[1]))
    return loss

def model(X_train, Y_train, X_test, Y_test, learning_rate=0.0001, num_epochs=1500, minibatch_size=32, print_cost=True):
    tf.reset_default_graph()
    tf.set_random_seed(1)
    seed = 3
    n_x, m = X_train.shape
    n_y = Y_train.shape[0]
    costs = []
    X, Y = create_placeholders(n_x, n_y)
    parameters = initialize_parameters(n_x, n_y)
    Z3 = forward_propagation(X, parameters)
    cost = compute_cost(Z3, Y)
    optimizer = tf.train.GradientDescentOptimizer(learning_rate=learning_rate).minimize(cost)
    init = tf.global_variables_initializer()
    start_time = time.time()
    with tf.Session() as sess:
        sess.run(init)
        for epoch in range(num_epochs):
            epoch_cost = 0.
            num_minibatches = m//minibatch_size
            seed = seed + 1
            minibatches = random_mini_batches(X_train, Y_train, minibatch_size, seed)

            for minibatch in minibatches:
                minibatch_X, minibatch_Y = minibatch
                _ , minibatch_cost = sess.run([optimizer, cost], feed_dict={X:minibatch_X, Y:minibatch_Y})
                epoch_cost += minibatch_cost / num_minibatches

            if print_cost == True and epoch % 10 == 0:
                costs.append(epoch_cost)
                end_time = time.time()
                print("Cost after epoch %i: %f,take %d seconds" % (epoch, epoch_cost,end_time-start_time))
        parameters = sess.run(parameters)
        print("Parameters have been trained!")
        # correct_prediction = tf.equal(Z3, Y)
        correct_prediction = abs(Z3 - Y) / Y < 0.2
        accuracy = tf.reduce_mean(tf.cast(correct_prediction, "float"))
        print("Train Accuracy:", accuracy.eval({X:X_train, Y:Y_train}))
        print("Test Accuracy:", accuracy.eval({X:X_test.T, Y:Y_test}))

        return parameters

# def my_model(x_vals_train,y_vals_train,x_vals_test,y_vals_test):
#     def init_weight(shape, st_dev):
#         weight = tf.Variable(tf.random_normal(shape, stddev=st_dev))
#         return weight
#
#     def init_bias(shape, st_dev):
#         bias = tf.Variable(tf.random_normal(shape, stddev=st_dev))
#         return bias
#
#     x_data = tf.placeholder(shape=[None, 14], dtype=tf.float32)
#     y_target = tf.placeholder(shape=[None, 1], dtype=tf.float32)
#
#     def fully_connected(input_layer, weights, biases):
#         layer = tf.add(tf.matmul(input_layer, weights), biases)
#         return tf.nn.relu(layer)
#
#     weight_1 = init_weight(shape=[7, 25], st_dev=10.0)
#     bias_1 = init_bias(shape=[25], st_dev=10.0)
#     layer_1 = fully_connected(x_data, weight_1, bias_1)
#
#     weight_2 = init_weight(shape=[25, 10], st_dev=10.0)
#     bias_2 = init_bias(shape=[10], st_dev=10.0)
#     layer_2 = fully_connected(layer_1, weight_2, bias_2)
#
#     weight_3 = init_weight(shape=[10, 3], st_dev=10.0)
#     bias_3 = init_bias(shape=[3], st_dev=10.0)
#     layer_3 = fully_connected(layer_2, weight_3, bias_3)
#
#     weight_4 = init_weight(shape=[3, 1], st_dev=10.0)
#     bias_4 = init_bias(shape=[1], st_dev=10.0)
#     final_output = fully_connected(layer_3, weight_4, bias_4)
#
#     loss = tf.reduce_mean(tf.reduce_sum(tf.square((y_target-final_output)), reduction_indices=[1]))
#     my_opt = tf.train.AdamOptimizer(0.01)
#     train_step = my_opt.minimize(loss)
#
#     # 填充数据与训练
#     init = tf.global_variables_initializer()
#     sess.run(init)



X_train_orig,Y_train_orig,X_test_orig,Y_test_orig = load_datasets()
X_train = X_train_orig.reshape(X_train_orig.shape[0],-1).T
X_test = X_test_orig.reshape(X_test_orig.shape[0],-1).T
Y_train = Y_train_orig
Y_test  = Y_test_orig
parameters = model(X_train,Y_train,X_test_orig,Y_test_orig,num_epochs=7000)
print(Y_test)
print(predict(X_test,parameters))#[[3.0283723 3.0283723 3.0283723 ... 3.0283723 3.0283723 3.0283723]]
for key,value in parameters.items():
    np.save("./save_parameter/%s.npy" % key,value)

# parameters = {}
# key_list = ["W1", "b1", "W2", "b2", "W3", "b3"]
# for key in key_list:
#     parameters[key] = np.load("./save_parameter/%s.npy" % key)
# print(predict(X_test,parameters))


