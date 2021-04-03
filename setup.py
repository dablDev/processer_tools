from setuptools import setup, find_namespace_packages

setup(
    name='processer_tools',
    packages=find_namespace_packages(include=['processer_tools*']),
    version='1.0',
    install_requires=[
        'pandas==0.24.2',
        'xlrd==1.2.0',
        'xlsxwriter==1.3.7'
    ],
    zip_safe=False
)
