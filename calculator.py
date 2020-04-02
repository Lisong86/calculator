from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog
from math import sqrt, asin, atanh, exp, pi, cosh, log, tanh
from matplotlib import pyplot as plt
from scipy.optimize import fsolve
import sys
from openpyxl import Workbook
import numpy as np
from ctypes import cdll
import os
from calculator_UI import Ui_MainWindow


class Mywindow(QMainWindow, Ui_MainWindow):
	def __init__(self):
		super(Mywindow, self).__init__()
		self.ui = Ui_MainWindow()
		self.ui.setupUi(self)
		self.ui.pushButton.clicked.connect(self.single_cal)
		self.ui.pushButton_2.clicked.connect(self.import_p)
		self.ui.pushButton_3.clicked.connect(self.export_p)
		self.ui.pushButton_4.clicked.connect(self.multi_cal)
		self.ui.pushButton_5.clicked.connect(self.open_excel)
		self.ui.pushButton_6.clicked.connect(self.figure)
		self.ui.radioButton.clicked.connect(self.calc_eq)
		self.ui.radioButton_2.clicked.connect(self.calc_manual)
		self.ui.radioButton_3.clicked.connect(self.beta)
		self.ui.radioButton_4.clicked.connect(self.beta)
		self.ui.radioButton_5.clicked.connect(self.beta)
		self.ui.radioButton_6.clicked.connect(self.beta)
		self.ui.radioButton_12.clicked.connect(self.setYRange)
		self.ui.radioButton_13.clicked.connect(self.setYRange)
		self.ui.radioButton_14.clicked.connect(self.setYRange)
		self.ui.radioButton_15.clicked.connect(self.setYRange)
		self.ui.radioButton_16.clicked.connect(self.setYRange)
		self.ui.radioButton_17.clicked.connect(self.setYRange)
		self.ui.radioButton_7.clicked.connect(self.elect11)
		self.ui.radioButton_8.clicked.connect(self.elect11)
		self.ui.radioButton_9.clicked.connect(self.elect21)

	def import_p(self):
		file=QFileDialog.getOpenFileName(self,"Open file",'./',("TXT Files(*.txt)"))
		if file[0]:
			with open(file[0],'r',errors='ignore') as f:
				txt=f.readlines()
			txt=[i.strip('\n') for i in txt]
			para={i:j for i,j in (t.split('=') for t in txt)}
			self.ui.lineEdit.setText(para['R'])
			self.ui.lineEdit_2.setText(para['T'])
			self.ui.lineEdit_3.setText(para['F'])
			self.ui.lineEdit_4.setText(para['CEC'])
			self.ui.lineEdit_5.setText(para['S'])
			self.ui.lineEdit_6.setText(para['x'])
			temp1,temp2=para['epsilon'].split('e')
			self.ui.lineEdit_7.setText(temp1)
			self.ui.lineEdit_8.setText(temp2)
			temp1,temp2=para['Aeff'].split('e')
			self.ui.lineEdit_9.setText(temp1)
			self.ui.lineEdit_10.setText(temp2)
			temp1,temp2=para['c'].split('e')
			self.ui.lineEdit_11.setText(temp1)
			self.ui.lineEdit_12.setText(temp2)
			self.ui.lineEdit_27.setText(para['start'])
			self.ui.lineEdit_28.setText(para['stop'])
			self.ui.lineEdit_29.setText(para['step'])
			self.ui.textEdit.setText(para['multi-c'])
			self.ui.textEdit_2.setText(para['multi-d'])

	def export_p(self):
		R = self.ui.lineEdit.text()
		T = self.ui.lineEdit_2.text()
		F =self.ui.lineEdit_3.text()
		CEC = self.ui.lineEdit_4.text()
		S = self.ui.lineEdit_5.text()
		epsilon1 = self.ui.lineEdit_7.text()
		epsilon2 = self.ui.lineEdit_8.text()
		A1 = self.ui.lineEdit_9.text()
		A2 =self.ui.lineEdit_10.text()
		c1=self.ui.lineEdit_11.text()
		c2=self.ui.lineEdit_12.text()
		d=self.ui.lineEdit_6.text()
		multi_c=self.ui.textEdit.toPlainText()
		multi_d=self.ui.textEdit_2.toPlainText()
		start = self.ui.lineEdit_27.text()
		stop = self.ui.lineEdit_28.text()
		step = self.ui.lineEdit_29.text()
		epsilon='%se%s'% (epsilon1,epsilon2)
		Aeff='%se%s'% (A1,A2)
		c='%se%s'% (c1,c2)
		key=['R','T','F','CEC','S','x','epsilon','Aeff','c','multi-c','start','stop','step','multi-d']
		value=[R,T,F,CEC,S,d,epsilon,Aeff,c,multi_c,start,stop,step,multi_d]
		s=''
		for i in range(14):
			s+='%s=%s'% (key[i],value[i])+'\n'
		file=QFileDialog.getSaveFileName(self,"Save file",'./',("TXT Files(*.txt)"))
		if file[0]:
			with open(file[0],'w',errors='ignore') as f:
				f.write(s)

	def calc_eq(self):
		self.ui.lineEdit_13.setEnabled(False)

	def calc_manual(self):
		self.ui.lineEdit_13.setText("-0.3")
		self.ui.lineEdit_13.setEnabled(True)

	def elect11(self):
		self.ui.radioButton_3.setEnabled(True)
		self.ui.radioButton_4.setEnabled(True)
		self.ui.radioButton_5.setEnabled(True)
		self.ui.radioButton_6.setEnabled(True)
		self.ui.lineEdit_14.setEnabled(True)
		self.beta()

	def setYRange(self):
		self.ui.lineEdit_32.setText("0")
		self.ui.lineEdit_33.setText("0")

	def elect21(self):
		self.ui.radioButton_3.setEnabled(False)
		self.ui.radioButton_4.setEnabled(False)
		self.ui.radioButton_5.setEnabled(False)
		self.ui.radioButton_6.setEnabled(False)
		self.ui.lineEdit_14.setText("2")
		self.ui.lineEdit_14.setEnabled(False)

	def beta(self):
		if self.ui.radioButton_3.isChecked():
			self.ui.lineEdit_14.setText("1")
		elif self.ui.radioButton_4.isChecked():
			self.ui.lineEdit_14.setText("1.11")
		elif self.ui.radioButton_5.isChecked():
			self.ui.lineEdit_14.setText("1.699")
		elif self.ui.radioButton_6.isChecked():
			self.ui.lineEdit_14.setText("2.506")

	def electrolyte(self):
		if self.ui.radioButton_7.isChecked():
			elect = 11
		elif self.ui.radioButton_8.isChecked():
			elect = 12
		else:
			elect = 21
		return elect

	def figureType(self):
		if self.ui.radioButton_12.isChecked():
			figure=7
			y_label='P(vdW) / atm'
		elif self.ui.radioButton_13.isChecked():
			figure=6
			y_label='P(edl) / atm'
		elif self.ui.radioButton_14.isChecked():
			figure=8
			y_label='P(hy) / atm'
		elif self.ui.radioButton_15.isChecked():
			figure=9
			y_label='P(dlvo) / atm'
		elif self.ui.radioButton_16.isChecked():
			figure=10
			y_label='P(net) / atm'
		else:
			figure=4
			y_label='E / V/m'
		return figure,y_label

	def concentration(self, elect):
		c = float(self.ui.lineEdit_11.text()) * 10 ** float(self.ui.lineEdit_12.text())
		if elect == 12:
			c *= 2
		return c

	def distance(self):
		d=float(self.ui.lineEdit_6.text())* 10 **-8
		return d

	def multi_distance(self):
		if self.ui.checkBox_2.isChecked():
			if self.ui.radioButton_10.isChecked():
				start = float(self.ui.lineEdit_27.text())
				stop = float(self.ui.lineEdit_28.text())
				step = float(self.ui.lineEdit_29.text())
				d = np.arange(start, stop, step) * 10 ** -8
			else:
				d = self.ui.textEdit_2.toPlainText()
				d = np.array(list(map(float, (d.split(","))))) * 10 ** -8
		else:
			d = np.array([self.distance()])
		return d

	def multi_concentration(self, elect):
		if self.ui.checkBox.isChecked():
			c = self.ui.textEdit.toPlainText()
			c = np.array(list(map(float, (c.split(",")))))
			if elect == 12:
				c = c * 2
		else:
			c = np.array([self.concentration(elect)])
		return c

	def get_value(self):
		self.R = float(self.ui.lineEdit.text())
		self.T = float(self.ui.lineEdit_2.text())
		self.F = float(self.ui.lineEdit_3.text())
		self.CEC = float(self.ui.lineEdit_4.text())
		self.S = float(self.ui.lineEdit_5.text())
		self.epsilon = float(self.ui.lineEdit_7.text()) * 10 ** float(self.ui.lineEdit_8.text())
		self.A = float(self.ui.lineEdit_9.text()) * 10 ** float(self.ui.lineEdit_10.text())
		self.Z = float(self.ui.lineEdit_14.text())
		self.elect = self.electrolyte()
		self.result_list = [
			['c', 'd', 'k', 'psi0', 'psimid', 'psix', 'E', 'sigma', 'pedl', 'pvdW', 'phy', 'pdlvo', 'pnet']]

	def calc(self, c, d):
		k = kappa(self.F, self.Z, c, self.epsilon, self.R, self.T)
		if self.ui.radioButton.isChecked():
			psi0 = surf_potential(self.CEC, self.S, self.Z, self.R, self.T, self.F, self.elect, k, c)
		else:
			psi0 = float(self.ui.lineEdit_13.text())
		psimid = mid_potential(self.Z, self.F, self.R, self.T, k, psi0, d)
		psix = x_potential(self.Z, self.F, self.R, self.T, self.epsilon, k, self.elect, psi0, c, d)
		E = electric_field(self.R, self.T, self.epsilon, self.Z, self.F, psix, c)
		sigma = charge_density(self.R, self.T, self.F, self.Z, self.elect, psi0, c, self.epsilon)
		pedl = force_edl(self.R, self.T, c, self.Z, self.F, psimid)
		pvdW = force_vdW(self.A, d)
		phy = force_hy(d)
		pdlvo = force_dlvo(pedl, pvdW)
		pnet = force_net(pedl, pvdW, phy)
		return k, psi0, psimid, psix, E, sigma, pedl, pvdW, phy, pdlvo, pnet

	def single_output(self, result):
		k, psi0, psimid, psix, E, sigma, pedl, pvdW, phy, pdlvo, pnet = result
		self.ui.lineEdit_15.setText('%.5f' % psi0)
		self.ui.lineEdit_16.setText('%.5f' % psimid)
		self.ui.lineEdit_17.setText('%.5f' % psix)
		self.ui.lineEdit_18.setText('%.3f' % sigma)
		self.ui.lineEdit_19.setText('%.2e' % k)
		self.ui.lineEdit_20.setText('%.2f' % (1 / k * 10 ** 9))
		self.ui.lineEdit_21.setText('%.2f' % pedl)
		self.ui.lineEdit_22.setText('%.2f' % pvdW)
		self.ui.lineEdit_23.setText('%.2f' % phy)
		self.ui.lineEdit_24.setText('%.2f' % pdlvo)
		self.ui.lineEdit_25.setText('%.2f' % pnet)
		self.ui.lineEdit_26.setText('%.4e' % E)

	def open_excel(self):
		if os.path.exists('data.xlsx'):
			os.startfile('data.xlsx')
		else:
			QMessageBox.information(self,'Notice','Data do not exist, please calculate.',QMessageBox.Yes,QMessageBox.Yes)

	def figure(self):
		self.get_value()
		x_min=float(self.ui.lineEdit_30.text())
		x_max=float(self.ui.lineEdit_31.text())
		y_min=float(self.ui.lineEdit_32.text())
		y_max=float(self.ui.lineEdit_33.text())
		d_list = np.arange(x_min,x_max,0.1)*10**-8
		c_list = self.multi_concentration(self.elect)
		len_c=len(c_list)
		result=[]
		for c in c_list:
			for d in d_list:
				result.append(self.calc(c, d))
		result=np.array(result)
		figureType=self.figureType()
		draw(y_min,y_max,figureType[1],len_c,d_list,c_list,result[:,figureType[0]])

	def single_cal(self):
		self.get_value()
		d = self.distance()
		c = self.concentration(self.elect)
		result = self.calc(c, d)
		self.single_output(result)

	def multi_cal(self):
		# if is_ExcelOpen('data.xlsx'):
		# 	QMessageBox.information(self,'Notice','Data file is open, please close.',QMessageBox.Yes,QMessageBox.Yes)
		# else:
		self.get_value()
		d_list = self.multi_distance()
		c_list = self.multi_concentration(self.elect)
		for c in c_list:
			for d in d_list:
				result = [c, d]
				result.extend(self.calc(c, d))
				self.result_list.append(result)
		save_excel(self.result_list)
		QMessageBox.information(self,'Notice','Calculation done!',QMessageBox.Yes,QMessageBox.Yes)


def kappa(F, Z, c, epsilon, R, T):
	k = sqrt(8 * pi * F * F * Z * Z * c / epsilon / R / T)
	return k


def surf_potential(CEC, S, Z, R, T, F, elect, k, c):
	if elect == 11:
		mean = CEC * k / S / Z * 10 ** -7
		a = fsolve(lambda x: 1 + 4 / (1 + x) - 4 / (1 + x * exp(-1)) - mean / c, -0.9999)[0]
		psi0 = -2 * R * T / Z / F * log((1 - a) / (1 + a))
	elif elect == 12:
		mean = CEC * k / S / Z * 10 ** -7
		b = fsolve(lambda x: 1 + 6 / (x - 1) - 6 / (x * exp(1) - 1) - mean / c, 1.0001)[0]
		psi0 = -1 * R * T / Z / F * log((b * b + 4 * b + 1) / (b * b - 2 * b + 1))
	else:
		mean = CEC * k / S / Z * 10 ** -7
		a1 = 2 + sqrt(3)
		a2 = 2 - sqrt(3)
		a3 = 3 * (2 + sqrt(3))
		a4 = 3 * (2 - sqrt(3))
		g = fsolve(
			lambda x: 1 - a3 / (exp(2 * x + 1) - a1) + a3 / (exp(2 * x) - a1) - a4 / (exp(2 * x + 1) - a2) + a4 / (
				exp(2 * x) - a2) - mean / c, 0.6585)[0]
		psi0 = -R * T / F * log((exp(2 * g) + 1) ** 2 / (exp(2 * g) ** 2 - 4 * exp(2 * g) + 1))
	return psi0


def mid_potential(Z, F, R, T, k, psi0, d):
	psimid = fsolve(lambda x: (pi / 2 * (
		1 + 0.25 * exp(2 * Z * F * x / R / T) + 1.3 * 1.3 / 2.4 / 2.4 * exp(4 * Z * F * x / R / T))) - (
		                          asin(exp(Z * F * (psi0 - x) / 2 / R / T))) - (
		                          0.25 * d * k * exp(-1 * Z * F * x / 2 / R / T)), psi0)[0]
	if psimid>=0:
		psimid=-0.00001
	return psimid


def x_potential(Z, F, R, T, epsilon, k, elect, psi0, c, d):
	if elect == 11:
		i = tanh(Z * F * psi0 / 4 / R / T)
		k = sqrt(8 * pi * F * F * Z * Z * c / epsilon / R / T)
		psix = 4 * R * T / Z / F * atanh(i * exp(-1 * k * d))
	elif elect == 12:
		i = (sqrt(1 + 2 * exp(-1 * F * psi0 / R / T)) + sqrt(3)) / (sqrt(1 + 2 * exp(-1 * F * psi0 / R / T)) - sqrt(3))
		k = sqrt(8 * pi * F * F * c / epsilon / R / T)
		psix = -1 * R * T / F * log(1 + 6 * i * exp(k * d) / (i * exp(k * d) - 1) ** 2)
	else:
		i = atanh(sqrt(1.0 / 3 * (1 + 2 * exp(F * psi0 / R / T))))
		k = sqrt(24 * pi * F * F * c / epsilon / R / T)
		psix = R * T / F * log(1.5 * (tanh(0.5 * k * d + i)) ** 2 - 0.5)
	return psix


def charge_density(R, T, F, Z, elect, psi0, c, epsilon):
	if elect == 11:
		c1 = c
		c2 = c
		Z1 = Z
		Z2 = -1 * Z
	elif elect == 12:
		c1 = 2 * c
		c2 = c
		Z1 = Z
		Z2 = -2 * Z
	else:
		c1 = c
		c2 = 2 * c
		Z1 = 2 * Z
		Z2 = -1 * Z
	sigma = sqrt(2 * epsilon * R * T / 4 / pi * (
		c1 * (exp(-1 * Z1 * F * psi0 / R / T) - 1) + c2 * (exp(-1 * Z2 * F * psi0 / R / T) - 1)))
	return sigma * 100


def electric_field(R, T, epsilon, Z, F, psix, c):
	E = -1 * sqrt(8 * pi * R * T / (epsilon * 10) * (c * 1000 * (exp(-1 * Z * F * psix / R / T) - 1)))
	return E


def force_edl(R, T, c, Z, F, psimid):
	force = 2 * R * T * c * (cosh(Z * F * psimid / R / T) - 1) / 100
	return force


def force_vdW(A, d):
	force = -1 * A / 6 / pi / (d / 10) ** 3 / 100000
	return force


def force_hy(d):
	force = 3.33 * exp(-5.76 * d / 10 * 10 ** 9) * 10 ** 4
	return force


def force_dlvo(edl, vdW):
	force = edl + vdW
	return force


def force_net(edl, vdW, hy):
	force = edl + vdW + hy
	return force


def save_excel(results):
	wb = Workbook()
	ws = wb.active
	for result in results:
		ws.append(result)
	wb.save('data.xlsx')

def draw(y_min,y_max,y_label,len_c,d_list,c_list,result_list):
	result=np.split(result_list,len_c,axis=0)
	# print(result)
	plt.cla()
	for i in range(len_c):
		x=d_list*10**8
		y=result[i]
		plt.plot(x,y)
	if y_min!=0 or y_max!=0:
		plt.ylim(y_min,y_max)
	plt.xlabel('d / nm')
	plt.ylabel(y_label)
	plt.legend(c_list)
	plt.show()

# def is_ExcelOpen(filename):
# 	_sopen = cdll.msvcrt._sopen
# 	_close = cdll.msvcrt._close
# 	_SH_DENYRW = 0x10
# 	if not os.access(filename, os.F_OK):
# 		return False # file doesn't exist
# 	h = _sopen(filename, 0, _SH_DENYRW, 0)
# 	if h == 3:
# 		_close(h)
# 		return False # file is not opened by anyone else
# 	return True # file is already open







if __name__ == "__main__":
	app = QApplication(sys.argv)
	w = Mywindow()
	w.show()
	sys.exit(app.exec_())
