import os

from setuptools import setup, find_packages

here = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(here, 'README.md')) as f:
    README = f.read()
with open(os.path.join(here, 'CHANGES.txt')) as f:
    CHANGES = f.read()

requires = []

setup(name='gdataext',
      version='0.1',
      description='Python Gdata Helpful Extensions',
      long_description=README + '\n\n' + CHANGES,
      classifiers=[
        "Programming Language :: Python"
        ],
      author='Nicholas Long',
      author_email='adoc@webmob.net',
      url='https://github.com/adoc/',
      keywords='gdata',
      packages=find_packages(),
      scripts=[],
      include_package_data=True,
      zip_safe=False,
      test_suite='gdataext.tests',
      install_requires=requires,
      test_requires=[]
      )
