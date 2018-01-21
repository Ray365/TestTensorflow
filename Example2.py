import tensorflow as tf
import numpy as np

'''
# create data
x_data = np.random.rand(1000).astype(np.float32)
y_data = x_data * 0.1 + 0.3

### create tensorflow structure start ###
Weights = tf.Variable(tf.random_uniform([1], -1.0, 1.0))
biases = tf.Variable(tf.zeros([1]))

y = Weights * x_data + biases

loss = tf.reduce_mean(tf.square((y - y_data)))

optimizer = tf.train.GradientDescentOptimizer(0.5)
train = optimizer.minimize(loss)

init = tf.global_variables_initializer()
#### create tensorflow structure edn ###

sess = tf.Session()
sess.run(init)  # Very important

for step in range(201):
    sess.run(train)
    if step % 20 == 0:
        print(step, sess.run(Weights), sess.run(biases))

# Variable 1
matrix1 = tf.constant([[2, 4]])
matrix2 = tf.constant([[3], [3]])

product = tf.matmul(matrix1, matrix2)

sess = tf.Session()
result = sess.run(product)
print('product result is {0}'.format(result))
sess.close()

# Variable 2
state = tf.Variable(0, name='counter')
one = tf.constant(1)

new_value = tf.add(state, one)
update = tf.assign(state, new_value)

init = tf.global_variables_initializer()  # must have if define variable

with tf.Session() as sess:
    sess.run(init)
    for _ in range(5):
        sess.run(update)
        print(sess.run(state))
'''

#Placeholder input value
input1=tf.placeholder(tf.float32)
input2=tf.placeholder(tf.float32)

output=tf.multiply(input1,input2)

with tf.Session() as sess:
    print(sess.run(output, feed_dict={input1:[7],input2:[2]}))