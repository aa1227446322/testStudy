"""
pytest生成测试报告
allure需要手动安装并添加环境变量
"""
import os
import pytest

if __name__ == '__main__':
    pytest.main(["-vs",
                "--capture=sys",  # 捕获输出
                # "test_framework.py",# 执行哪个文件 当前文件名
                "test/",  # 执行test/下所有文件
                "--clean-alluredir",  # 执行前清除掉上次执行的数据#
                "--alluredir=allure-result"  # 本次执行结果数据文件夹
    ])
    print("用例执行完毕")
    # os就是系统操作命令   在windows下等同于命令台
    os.system("allure generate allure-result -o ./report_allure --clean")