import setuptools #导入setuptools打包工具
 
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()
 
setuptools.setup(
    name="excel_operate-light22", # 用自己的名替换其中的YOUR_USERNAME_
    version="1.0.2",    #包版本号，便于维护版本
    author="北极星光",    #作者，可以写自己的姓名
    author_email="light22@126.com",    #作者联系方式，可写自己的邮箱地址
    license='MIT',
    description="方便快捷Excel操作",#包的简述
    long_description=long_description,    #包的详细介绍，一般在README.md文件内
    long_description_content_type="text/markdown",
    url="https://github.com/18513233125/Excel-operate",    #自己项目地址，比如github的项目地址
    download_url = 'https://pypi.org/project/excel-operate-light22',
    packages=setuptools.find_packages(),
    install_requires=['openpyxl'],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',    #对python的最低版本要求
)