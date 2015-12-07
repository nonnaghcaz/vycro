#! /usr/bin/env python

from __future__ import absolute_import, print_function, division
try:
    from setuptools import setup, find_packages
    from setuptools.command.test import test as TestCommand
except ImportError:
    raise RuntimeError('No suitable version of setuptools detected.')
import re
import os
import codecs
import sys

HERE = os.path.abspath(os.path.dirname(__file__))
# sys.path.insert(0, os.path.join(HERE, 'vycro'))

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
    tests_require=['pytest', 'pytest-pep8', 'pytest-cov'],
    test_suite='vycro.tests.test_vycro',
    cmdclass={'test': PyTest}
)
