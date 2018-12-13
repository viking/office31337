from setuptools import setup

setup(name='office31337',
      version='0.1',
      description='Utilities to fetch e-mails from Office 365',
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
