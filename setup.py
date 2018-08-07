import setuptools

with open('README.md', 'r') as fh:
    long_description = fh.read()


setuptools.setup(
    name='xlsxmetadata',
    version='0.0.3',
    author='spmassot',
    author_email='spmassot@gmail.com',
    description='Really lightweight lib for peeking into xlsx column/row size before you try to open the file with something else',
    long_description=long_description,
    url='https://github.com/spmassot/xlsxmetadata',
    packages=['xlsxmetadata'],
    zip_safe=False
)
