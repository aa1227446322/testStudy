from string import  Template
"""
变量渲染:Template用法
字符串里被 括起来的字符，与 字典中的 key一致! 直接用value 替换掉 这个字符
"""
dc = {"teacher":"小明"}
st ="欢迎大家来到，${teacher}的课堂"
print(st)
print(Template(st).substitute(dc))







