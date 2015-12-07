#! /usr/bin/env python

from __future__ import absolute_import, print_function, division
try:
    from setuptools import setup, find_packages
    from setuptools.command import easy_install
    from setuptools.command.test import test as TestCommand
except ImportError:
    raise RuntimeError('No suitable version of setuptools detected.')
import re
import os
import codecs
import sys
import struct

try:
    import pip
except ImportError:
    pip = None

HERE = os.path.abspath(os.path.dirname(__file__))
# sys.path.insert(0, os.path.join(HERE, 'vycro'))

FEDEX_LINKS = (
    '\\\\fxg\\dfs\\SHARED\\GROUNDCOMM\\Gannon'
    '\\GPS Navigation\\Garmin\\Reporting\\deps'
)
DEPENDENCY_LINKS = []

PYVER = sys.version_info
PYPLAT = sys.platform
PY2_MIN = (2, 7)
PY2_MAX = (2, 7)
PY3_MIN = (3, 4)
PY3_MAX = (3, 4)

###############################################################################


class PyTest(TestCommand):
    def finalize_options(self):
        TestCommand.finalize_options(self)
        self.test_args = ['--strict', '--verbose', '--tb=long', 'tests']
        self.test_suite = True

    def run_tests(self):
        import pytest
        errcode = pytest.main(self.test_args)
        sys.exit(errcode)

###############################################################################

whl = os.path.join(FEDEX_LINKS, 'whl')
if os.path.exists(whl):
    DEPENDENCY_LINKS.append(whl)


def is_python_64bit():
    return (struct.calcsize("P") == 8)


def read(*parts):
    with codecs.open(os.path.join(HERE, *parts), 'rb', 'utf-8') as f:
        return f.read()

META_FILE = read(*['vycro', '__about__.py'])


def find_meta(meta):
    meta_match = re.search(
        r"^__{meta}__ = ['\"]([^'\"]*)['\"]".format(meta=meta),
        META_FILE, re.M
    )

    if meta_match:
        return meta_match.group(1)
    raise RuntimeError("Unable to find __{meta}__ string.".format(meta=meta))


def install_pywin32():
    print('\nAttempting to install Pywin32 dependency locally...')
    if PYPLAT is 'win32':
        if (
            (PYVER >= PY3_MIN and PYVER <= PY3_MAX) or
            (PYVER >= PY2_MIN and PYVER <= PY2_MAX)
        ):
            (major, minor) = PYVER

            if is_python_64bit():
                pybit = 'win_amd64'
            else:
                pybit = 'win32'

            if pip is not None:
                dep = ''.join([
                    'pypiwin32-219-cp', str(major), str(minor), '-none-',
                    pybit, '.whl'
                ])
                pip.main([
                    'install', os.path.join(FEDEX_LINKS, 'whl', dep)
                ])
            else:

                dep = ''.join([
                    'pywin32-219.', str(pybit).replace('_', '-'),
                    '-py', str(major), '.', str(minor), '.exe'
                ])
                easy_install.main([
                    '-U', os.path.join(FEDEX_LINKS, 'bin', dep)
                ])
        else:
            raise RuntimeError('No suitable python version for pywin32')
    else:
        raise RuntimeError(
            'No suitable OS found.'
            'Only 32 or 64 bit versions of Microsoft Windows valid.'
        )

# if 'install' in sys.argv:
#     install_pywin32()

setup(
    name='vycro',
    version=find_meta('version'),
    url=find_meta('uri'),
    license=find_meta('license'),
    author=find_meta('author'),
    author_email=find_meta('email'),
    description=find_meta('summary'),
    maintainer=find_meta('author'),
    maintainer_email=find_meta('email'),
    keywords=['VBA', 'macros', 'Excel'],
    long_description=read('README.rst'),
    packages=find_packages(exclude=['contrib', 'docs', 'tests']),
    zip_safe=False,
    classifiers=[
        'Natural Language :: English',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Topic :: Software Development :: Libraries :: Python Modules'
    ],
    install_requires=['pypiwin32'],
    include_package_data=True,
    dependency_links=DEPENDENCY_LINKS,
    tests_require=['pytest', 'pytest-pep8', 'pytest-cov'],
    test_suite='vycro.tests.test_vycro',
    cmdclass={'test': PyTest}
)
