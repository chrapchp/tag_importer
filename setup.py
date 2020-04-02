from setuptools import setup, find_packages

setup(name='tag_importer', 
      version='0.1.0', 
      packages=[], 
      install_requires=[ 'pandas', 'xlrd','click' ],
      description='CRUD functionality to manage tags in Ovarro TBOX RTU',
      author='pjc',
      url='https://github.com/chrapchp/tag_importer', 
      entry_points={
        'console_scripts': ['tag_importer = tag_importer.cli:start'] }
      )