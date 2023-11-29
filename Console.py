"""  
@Time: 2023/11/29 11:02 
@Auth: Y5neKO
@File: Console.py 
@IDE: PyCharm 
"""
import argparse
import sys

from HighRiskVul import *
from HighRiskPort import *

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="使用帮助")
    scanner = parser.add_argument_group('处理参数')
    scanner.add_argument("-e", type=str, dest="type", default="vul", choices=["vul", "port"],
                         help="指定操作类型, 默认为高危漏洞。vul: 高危漏洞 | port: 高危端口")
    scanner.add_argument("--input", type=str, dest="input", help="要处理的表格")
    scanner.add_argument("--output", type=str, dest="output", help="要导出的表格")
    # scanner.add_argument("-sl", type=)

    args = parser.parse_args()

    if args.input is None or args.output is None:
        print("请输入参数！\n使用-h查看帮助")
        sys.exit(1)
    if args.type == "vul":
        vul_main(args.input, args.output)
    elif args.type == "port":
        port_main()
