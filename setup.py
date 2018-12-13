from setuptools import setup

def readme():
    with open('README.md') as f:
        return f.read()

setup(name='office31337',
      version='0.1',
      description='Utility to fetch e-mails from Office 365',
      long_description=readme(),
      url='https://github.com/viking/office31337',
      author='Jeremy Stephens',
      author_email='jeremy.f.stephens@vumc.org',
      license='MIT',
      scripts=['bin/office31337'],
      packages=['office31337'],
      install_requires=[
          'exchangelib',
          'unidecode',
          'keyring'
      ],
      zip_safe=False)
