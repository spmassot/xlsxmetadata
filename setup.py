import setuptools

with open('README.md', 'r') as fh:
    long_description = fh.read()


setuptools.setup(
    name='xlsxmetadata',
    version='0.0.2',
    author='automatonymous',
    author_email='spmassot@gmail.com',
    description='Really lightweight lib for peeking into xlsx column/row size before you try to open the file with something else',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/automatonymous/xlsxmetadata',
    packages=['xlsxmetadata'],
    zip_safe=False
)
