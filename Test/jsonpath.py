# 对excel的提取参数做判断，如果有值，应的参数提取出来。如果没值，就不提取
import  jsonpath
case_info={}
rs = {}
dic = {}
if case_info["提取参数"]:
    # 有值，所以开始提取--我怎么知道要提取的值在哪一层?\
    # jsonpath 提取
    lk = jsonpath.jsonpath(rs.json(),"$.."+case_info["提取参数"])
    dic[case_info["提取参数"]]= lk[0]