from setuptools import setup
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.rst')) as f:
    long_description = f.read()
version = '1.2.0'

setup(
    name='nanonispy2',
    version=version,
    description='Library to parse Nanonis files. Forked from https://github.com/underchemist/nanonispy',
    long_description=long_description,
    url='https://github.com/ceds92/nanonispy2',
    author='Julian Ceddia',
    author_email='jdceddia@gmail.com',
    license='MIT',
    platforms=['any'],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Science/Research',
        'Topic :: Scientific/Engineering :: Physics',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9'
    ],
    keywords='science numpy parse library',
    python_requires='>=3.6',
    packages=['nanonispy2'],
    package_data={'nanonispy2': ['LICENSE', 'README.md'], },
    install_requires=['numpy', ],
    tests_require=['nose', 'coverage', ],
    include_package_data=True,
)
