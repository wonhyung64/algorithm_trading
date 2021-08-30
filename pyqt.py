#%%
import sys
from PyQt5.QtWidgets import *

app = QApplication(sys.argv)
print(sys.argv)
label = QLabel("Hello PyQt")
label.show()
app.exec_()
# %%
import sys
from PyQt5.QtWidgets import *

app = QApplication(sys.argv)
label = QPushButton("Quit")
label.show()
app.exec_()

# %%
