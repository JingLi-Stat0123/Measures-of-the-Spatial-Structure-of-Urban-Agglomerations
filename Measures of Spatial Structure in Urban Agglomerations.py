import numpy as np
import pandas as pd
data = pd.read_excel(r"D:\HuaweiMoveData\Users\HUAWEI\Desktop\data_多sheet表格式.xlsx",sheet_name=None)   #sheet_name默认是0读取第一张sheet表
city_mapping = {
    "京津冀": ["北京市", "天津市", "石家庄市", "唐山市", "秦皇岛市", "保定市", "张家口市", "承德市", "沧州市", "廊坊市"],
    "太原": ["太原市", "晋中市", "忻州市", "吕梁市"],
    "呼包鄂榆": ["呼和浩特市", "包头市", "鄂尔多斯市", "榆林市"],
    "辽中南": ["沈阳市", "大连市", "鞍山市", "抚顺市", "本溪市", "丹东市", "营口市", "辽阳市", "盘锦市", "铁岭市"],
    "哈长": ["长春市", "吉林市", "松原市", "哈尔滨市", "齐齐哈尔市", "大庆市", "牡丹江市"],
    "长三角": ["上海市", "南京市", "无锡市", "徐州市", "常州市", "苏州市", "南通市", "连云港市", "淮安市", "盐城市", "扬州市", "镇江市", "泰州市", "宿迁市", "杭州市", "宁波市", "温州市", "嘉兴市", "湖州市", "绍兴市", "金华市", "衢州市", "舟山市", "台州市", "丽水市"],
    "江淮": ["合肥市", "芜湖市", "马鞍山市", "铜陵市", "安庆市", "滁州市", "六安市", "池州市", "宣城市"],
    "海峡西岸": ["福州市", "厦门市", "莆田市", "三明市", "泉州市", "漳州市", "南平市", "龙岩市", "宁德市"],
    "环鄱阳湖": ["南昌市", "景德镇市", "九江市", "新余市", "鹰潭市", "吉安市", "宜春市", "抚州市", "上饶市"],
    "山东半岛": ["济南市", "青岛市", "淄博市", "东营市", "烟台市", "潍坊市", "威海市", "日照市"],
    "中原": ["郑州市", "开封市", "洛阳市", "平顶山市", "新乡市", "焦作市", "许昌市", "漯河市"],
    "武汉": ["武汉市", "黄石市", "鄂州市", "孝感市", "黄冈市", "咸宁市"],
    "长株潭": ["长沙市", "株洲市", "湘潭市", "衡阳市", "岳阳市", "常德市", "益阳市", "娄底市"],
    "珠三角": ["广州市", "深圳市", "珠海市", "佛山市", "江门市", "肇庆市", "惠州市", "东莞市", "中山市"],
    "北部湾": ["南宁市", "北海市", "防城港市"],
    "成渝": ["重庆市", "成都市", "自贡市", "泸州市", "德阳市", "绵阳市", "遂宁市", "内江市", "乐山市", "南充市", "眉山市", "宜宾市", "达州市", "雅安市", "资阳市"],
    "黔中": ["贵阳市", "遵义市", "安顺市"],
    "滇中": ["昆明市", "曲靖市", "玉溪市"],
    "关中-天水": ["西安市", "铜川市", "宝鸡市", "咸阳市", "渭南市", "商洛市", "天水市"],
    "兰州-西宁": ["兰州市", "白银市", "定西市", "西宁市"],
    "宁夏沿黄": ["石嘴山市", "吴忠市", "中卫市"]
}
def calculate_q(values,max_value):   # values表示排序之后的数组
    for i in np.arange(1,len(values)):  # 跳过第一个
        rank=i+1
        q_i=(np.log(max_value/float(values[i])))/np.log(rank)
        q_sum=+q_i
        count=+1
    return q_sum/count
result = []
for year,df_year in data.items():
    for group,city_list in city_mapping.items():
        group_data=df_year[df_year["城市"].isin(city_list)].copy()
        group_data_pop=group_data.sort_values(by="人口",ascending=False) #虽然排序了但本质是全部数据
        max_value=group_data_pop["人口"].iloc[0]
        pop_val=group_data_pop["人口"].values
        q_M=calculate_q(pop_val,max_value)
        # 计算经济维度
        group_data_gdp=group_data.sort_values(by="GDP",ascending=False)
        max_val_gdp=group_data_gdp["GDP"].iloc[0]
        gdp_val= group_data_gdp["GDP"].values
        q_N=calculate_q(gdp_val,max_val_gdp)
        index_q=q_M*q_N
        result.append({
        "城市群":group,
        "q_M":round(q_M,3),
        "q_N":round(q_N,3),
        "总指数":round(index_q,4),
        "年份":year})
final_df=pd.DataFrame(result)
output_path=r"D:\HuaweiMoveData\Users\HUAWEI\Desktop\城市群空间结构最终指标.xlsx"
final_df.to_excel(output_path)
print(f"最终计算的结果保存至：{output_path}")
