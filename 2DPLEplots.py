import numpy as np
import matplotlib.pyplot as plt

filename=input("name of the file:")
data = np.genfromtxt(filename, unpack=True, delimiter='\t')

x=data[0,2:]
y=data[1:,0]
xmin=x[0]
xmax=x[-1]
ymin=y[0]
ymax=y[-1]
values = data[1:,2:]
values2=np.rot90(values.T)
fig = plt.figure(figsize=(8,8))
ax = fig.gca()
ax.set_xlim(xmin, xmax)
ax.set_ylim(ymin, ymax)
ax.imshow(values2, cmap='coolwarm', extent=[xmin, xmax, ymin, ymax])
ax.set_xlabel('emission wavelength [nm]')
ax.set_ylabel('excitation wavelength [nm]')
plt.title('PLE map for ' + filename)
plt.show()