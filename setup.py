from setuptools import setup, find_packages

setup(
	name='pdfcraft',
	version='0.1.0',
	packages=find_packages(),
	install_requires=['docxtpl', 'docx2pdf'],
	author='mrchn',
	description='lib for automate pdf generation from .docx templates',
	python_requires='>=3.8',
)