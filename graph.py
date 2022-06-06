import matplotlib.pyplot as plt
import numpy as np

labels = ['Fase U/V', 'Fase V/W', 'Fase W/U', 'Fase U/M', 'FaseV/M', 'Fase W/M']
men_means = [1000, 1233, 1212, 3333, 4443, 2323]
women_means = [11123, 22234, 33424, 22423, 23323, 44423]

x = np.arange(len(labels))  # the label locations
width = 0.35  # the width of the bars

fig, ax = plt.subplots()
rects1 = ax.bar(x - width/2, men_means, width, label='1 min')
rects2 = ax.bar(x + width/2, women_means, width, label='10 min')

# Add some text for labels, title and custom x-axis tick labels, etc.
ax.set_ylabel('Resistência (MegaOhm)')
#ax.set_title('Título')
ax.set_xticks(x, labels)
ax.legend()

ax.bar_label(rects1, padding=3)
ax.bar_label(rects2, padding=3)

fig.tight_layout()

plt.savefig('graphs\\stern')