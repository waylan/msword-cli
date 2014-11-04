from setuptools import setup

with  open('README.rst', mode='r') as fd:
    long_description = fd.read()

setup(
    name='MSWord-CLI',
    version='0.1',
    url='https://github.com/waylan/msword-cli/',
    description='A Command line interface for MS Word.',
    long_description=long_description,
    author='Waylan Limberg',
    author_email='waylan.limberg@icloud.com',
    license='BSD License',
    py_modules=['msword_cli'],
    install_requires=[
        'pywin32',
        'click>=3'
    ],
    entry_points='''
        [console_scripts]
        msw=msword_cli:cli
    ''',
    test_suite = 'tests',
    tests_require =['mock'],
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Environment :: Console',
        'Environment :: Win32 (MS Windows)',
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Developers',
        'Intended Audience :: System Administrators',
        'License :: OSI Approved :: BSD License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Topic :: Office/Business',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Utilities'
    ]
)