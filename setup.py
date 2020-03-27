# -*- coding: utf-8 -*-

from setuptools import setup


version = '1.0.4'


setup(
    name='django-excel-base',
    version=version,
    keywords='Django Excel Base',
    description='Django Excel Base',
    long_description=open('README.rst').read(),

    url='https://github.com/django-xxx/django-excel-base',

    author='Hackathon',
    author_email='kimi.huang@brightcells.com',

    packages=['django_excel_base'],
    py_modules=[],
    install_requires=['xlwt', 'pytz', 'screen'],

    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: Web Environment',
        'Framework :: Django',
        'Intended Audience :: Developers',
        'Programming Language :: Python',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
    ],
)
