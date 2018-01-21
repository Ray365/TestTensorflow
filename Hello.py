#2018-1-14 8:20
import tensorflow as tf

state = tf.Variable(0, name='counter')

one = tf.constant(1)

new_value = tf.add(state, one)

print(new_value)

update = tf.assign(state, new_value)
print(update)
