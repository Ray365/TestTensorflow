import tensorflow as tf
import numpy as np

# create data
x_data = np.random.rand(100).astype(np.float32)
print('x_data:', x_data)

y_data = x_data * 0.1 + 0.3
print('y_data:', y_data)
