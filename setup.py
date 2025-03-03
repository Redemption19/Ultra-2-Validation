from setuptools import setup, find_packages

setup(
    name="ultra2_validation_dash",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        'openpyxl==3.1.2',
        'streamlit==1.24.0',
        'pandas==2.0.3',
        'pathlib==1.0.1',
        'Pillow==10.0.0',
        'rich==13.9.4'
    ],
    python_requires='>=3.8,<3.13'
) 