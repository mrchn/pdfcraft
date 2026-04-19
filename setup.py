from setuptools import setup, find_packages
setup(
	name='pdfcraft-mrchn', version='0.1.4', author='mrchn',
	packages=find_packages(), install_requires=['docxtpl', 'docx2pdf'],
	description='lib for automate pdf generation from .docx templates', python_requires='>=3.8'
)