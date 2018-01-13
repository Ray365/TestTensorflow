import tensorflow as tf

a = 2
b = 3
x = tf.add(a, b)
y = tf.multiply(a, b)
useless = tf.multiply(a, x)
z = tf.pow(y, x)
with tf.Session() as sess:
    writer = tf.summary.FileWriter("./graphs", sess.graph)
    z = sess.run(z)
    writer.close()
