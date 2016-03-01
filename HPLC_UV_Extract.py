
from win32com.client import Dispatch;
from win32com.client import VARIANT as variant;
from pythoncom import *;
import sys;
import matplotlib.pyplot as plt;
import numpy as np;
from win32com.client.gencache import EnsureDispatch;
import time;
import datetime;
import os;
from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm

fileName = sys.argv[2];
flagString = sys.argv[1];

if flagString == 'preview':
	outputFlag = 0;
elif flagStirng == 'export':
	outputFlag = 1;
else:
	print('Unknow Parameter. \nPossible Flags: preview, export\nOperating in preview mode.');
	outputFlag = 0;
	
try:
	displayIndex0 = int(sys.argv[3]);
	displayIndex1 =	int(sys.argv[4]);
except IndexError:
	print('Display index range not given!\nUsing default: 1-200');
	displayIndex0 = 1;
	displayIndex1 = 200;

dummyVariant_arr = variant(VT_BYREF, []);
dummyVariant_val = variant(VT_BYREF, 0);

obj = Dispatch('MSFileReader.XRawFIle');
print(str(obj))
obj.Open(fileName);
print('File Loaded: ' + os.path.basename(fileName));
obj.SetCurrentController(3, 1);
print('Detector Controller Set: PDA');
numSpec = obj.GetNumSpectra(0);
print(str(numSpec) + ' spectra found.');
print('Display time range: ' + str(obj.RTFromScanNum(displayIndex0, 0)) + ' - ' + str(obj.RTFromScanNum(displayIndex1, 0)));
print('Loading data:');
waveLength = [];
intensity = [];
timeaxis = [];
for C1 in range(displayIndex0, displayIndex1):
	temp = obj.GetMassListFromScanNum(C1, '', 0, 0, 0, 0,dummyVariant_val.value, dummyVariant_arr, dummyVariant_arr, dummyVariant_val.value);
	pdrt = obj.RTFromScanNum(C1, 0);
	waveLength.append(temp[2][0]);
	intensity.append(temp[2][1]);
	timeaxis.append([pdrt for x in range(0, len(temp[2][0]))]);
	sys.stdout.write('\r' + str(C1 - displayIndex0) + '/' + str(displayIndex1 - displayIndex0));
print('Data Loaded');
	
fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')
ax.plot_surface(waveLength, timeaxis, intensity, rstride=10, cstride=10, cmap = cm.afmhot);
ax.set_xlabel('Wavelength (nm)');
ax.set_ylabel('Time (mintues)');
ax.set_zlabel('a.u.');
plt.show()