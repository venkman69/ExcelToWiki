'''
Created on Jul 28, 2015

@author: venkman69@yahoo.com
'''
from setuptools import setup

setup(name='exceltowiki',
      version='0.1.11',
      description='Convert Excel to Wiki while maintaining formatting',
      url='http://github.com/venkman69/ExcelToWiki',
      author='Narayan Natarajan',
      author_email='venkman69@yahoo.com',
      license='MIT',
      packages=['exceltowiki'],
      install_requires=[
          'openpyxl',
      ],
      zip_safe=False)
