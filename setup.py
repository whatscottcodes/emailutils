from setuptools import setup, find_packages

setup(name='emailutils',
      version='0.1',
      py_modules=["emailutils"],
      description='email utils',
      long_description='functions to download file from outlook or send emails - requires outlook',
      classifiers=[
          'Development Status :: 3 - Alpha',
          'License :: OSI Approved :: MIT License',
          'Programming Language :: Python :: 3.7',
          'Topic :: Email Automation :: Emails',
      ],
      keywords='email some radical files',
      url='http://github.com/whatscottcodes',
      author='SNelson',
      author_email='scott.nelsonjr@gmail.com',
      license='MIT',
      packages=find_packages(),
      install_requires=[
          'markdown', 'pywin32'
      ],
      include_package_data=False,
      zip_safe=False)