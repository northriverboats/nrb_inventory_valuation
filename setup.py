from setuptools import setup

setup(
    name='ivs',
    version='0.1',
    py_modules=['ivs'],
    install_requires=[
        'Click',
    ],
    entry_points='''
        [console_scripts]
        ivs=ivs:cli
    ''',
)
