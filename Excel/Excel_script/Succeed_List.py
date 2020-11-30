# 导入包
import pandas as pd
import xlwings as xw
import numpy as np
# 操作思维逻辑
# 应用->工作簿->工作表->范围
app = xw.App(add_book=False,visible=False)
# wb = app.books.add()

wb = app.books.open("../Excel_Test/10月汉滨区住宅成交数据.xlsx")
Ple_li = app.books.open('../Excel_Test/板块对应表.xlsx')
sht1 = wb.sheets['普通住宅']

class make_list:
    def __init__(self):
        self.data = pd.read_excel('../Excel_Test/10月汉滨区住宅成交数据.xlsx', sheet_name="普通住宅")

        #项目名称
        self.project_Name = sht1.range("b2:b80").value
        # print(project_Name)

    def proje(self):
        Project_Name = []
        for i in self.project_Name:
            if i not in Project_Name:
                Project_Name.append(i)

    # 获取套数

        Building = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            # print(some)
            # 指定套数列求和
            Sum = some['套数'].sum()
            Building.append(Sum)

    # 获取面积

        Area = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            # print(some)
            # 指定套数列求和
            Sum = some['面积'].sum()
            # print(Sum)
            Area.append(Sum)

    #获取总套数

        Tower_all = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            # print(some)
            # 指定套数列求和
            Sum = some['总套数'].sum()
            Tower_all.append(Sum)


    #获取总面积

        Area_all = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            # print(some)
            # 指定套数列求和
            Sum = some['总面积'].sum()
            Area_all.append(Sum)


    #获取可售套数

        Sell_number = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            # print(some)
            # 指定套数列求和
            Sum = some['可售套数'].sum()
            Sell_number.append(Sum)



    #获取可售面积
        Sell_area = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['可售面积'].sum()
            Sell_area.append(Sum)


    #获取已售套数
        Already_tower_all = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['已售套数'].sum()
            Already_tower_all.append(Sum)
    #获取已售面积
        already_area_al = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['已售面积'].sum()
            already_area_al.append(Sum)
    #获取不可售套数
        Not_for_sale = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['不可售套数'].sum()
            Not_for_sale.append(Sum)

    #获取不可售面积
        already_area_all = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['不可售面积'].sum()
            already_area_all.append(Sum)

    #获取最高价
        tall_price = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['最高价'].max()
            tall_price.append(Sum)

    #获取最低价
        low_price = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            Sum = some['最低价'].min()
            low_price.append(Sum)

    #获取平均值
        average = []
        for i in Project_Name:
            some = self.data[(self.data['QubuilderName'] == i)]
            #均价
            aver = some['均价']
            area = some['面积']
            sum = some['面积'].sum()
            sum_1 = 0
            for Aver,Area_1 in zip(aver,area):
                sum_1 += Aver*Area_1
            result = sum_1/sum
            result = round(result,2)
            #平均值我这里保留的两位小数可以调整round的参数进行调整
            average.append(result)


        return Project_Name,Building, Area, Tower_all,Area_all, Sell_number,Sell_area,Already_tower_all,Not_for_sale,already_area_al,tall_price,low_price,average

if __name__ == '__main__':
    make = make_list()
    Project_Name,Building, Area, Tower_all,Area_all, Sell_number,Sell_area,Already_tower_all,Not_for_sale,already_area_al,tall_price,low_price,average = make.proje()
    sht_two = Ple_li.sheets['Sheet1']
    # 读取板块表的项目名称以及板块  pt_name_two ---> project_name  plates ----> plate
    pt_name_two = sht_two.range("a2:a93").value
    pt_name_two[75] = '安康吾悦广场商住项目'
    pt_name_two[1] = '阳光尚都'
    pt_name_two[8] = "御公馆三期"
    pt_name_two[19] = "长兴小区三期"
    pt_name_two[62] = "清水园新型社区"
    pt_name_two[11] = "南龙滨江公馆"
    pt_name_two[86] = "新强怡景苑"
    pt_name_two[28] = "广景花园三期"
    pt_name_two[24] = "安康御华城（东盛.澜悦湾)"
    pt_name_two[53] = "康兴园小区3号楼"
    pt_name_two[87] = "御景华府小区"
    pt_name_two[56] = "美康华庭小区"
    pt_name_two[20] = "城市风景二期"




    plates = sht_two.range("b2:b93").value
    # print(pt_name_two,plates)
    dic = {}
    for pt_names, plate in zip(pt_name_two, plates):
        if pt_names in Project_Name:
            dic[pt_names] = plate


    one_list = []
    for i in Project_Name:
        if i in dic.keys():
            one_list.append(dic[i])

    wb = app.books.add()
    # 工作表
    sht = wb.sheets["sheet1"]
    sht.range("a1").value = ["月度","物业类型","物业细分","区域","板块","项目名称","套数","面积","最高价",
                             "最低价","均价","总套数","总面积","可销售套数","可销售面积","已销售套数","已销售面积","备注"]
    _date_time = "2020/10/1"
    sht.range("a2:a24").options(transpose=True).value = [_date_time for i in range(31)]
    _type = "住宅"
    sht.range("b2:b24").options(transpose=True).value = [_type for i in range(31)]
    _careful = "普通住宅"
    sht.range("c2:c24").options(transpose=True).value = [_careful for i in range(31)]
    _site = "汉滨区"
    sht.range("d2:d24").options(transpose=True).value = [_site for i in range(31)]

    sht.range("f2:f24").options(transpose=True).value = Project_Name
    sht.range("g2:g24").options(transpose=True).value = Building
    #面积
    sht.range("h2:h24").options(transpose=True).value = Area
    sht.range("i2:i24").options(transpose=True).value = tall_price
    #最低价
    sht.range("j2:j24").options(transpose=True).value = low_price
    #均价
    sht.range("k2:k24").options(transpose=True).value = average
    #总套数
    sht.range("l2:l24").options(transpose=True).value = Tower_all
    #总面积
    sht.range("m2:m24").options(transpose=True).value = Area_all
    #可销售套数
    sht.range("n2:n24").options(transpose=True).value = Sell_number
    #可销售面积
    sht.range("o2:o24").options(transpose=True).value = Sell_area
    #已销售套数
    sht.range("p2:p24").options(transpose=True).value = Already_tower_all
    #可销售面积
    sht.range("q2:q24").options(transpose=True).value = already_area_al
    sht.range("e2:e24").options(transpose=True).value = one_list

    da_1 = sht.range("a4:q4").value
    da_2 = sht.range("a11:q11").value
    da_3 = sht.range("a12:q12").value
    da_4 = sht.range("a16:q16").value
    da_5 = sht.range("a19:q19").value
    de = sht.range("A4,A11,A12,A16,A19").api.EntireRow.Delete()
    # print(da_1,da_2,da_3,da_4,da_5)


    wb_1 = app.books.add()
    sht_1 = wb_1.sheets["sheet1"]
    sht_1.range("a1").value = ["月度", "物业类型", "物业细分", "区域", "板块", "项目名称", "套数", "面积", "最高价",
                             "最低价", "均价", "总套数", "总面积", "可销售套数", "可销售面积", "已销售套数", "已销售面积", "备注"]
    sht_1.range("a2").value = da_1
    sht_1.range("a3").value = da_2
    sht_1.range("a4").value = da_3
    sht_1.range("a5").value = da_4
    sht_1.range("a6").value = da_5
    _site_1 = "高新区"
    sht_1.range("d2:d37").options(transpose=True).value = [_site_1 for i in range(30)]
    _date_time_1 = "2020/10/1"
    sht_1.range("a7:a32").options(transpose=True).value = [_date_time_1 for i in range(25)]
    _type_1 = "住宅"
    sht_1.range("b7:b32").options(transpose=True).value = [_type_1 for i in range(25)]
    _careful_1 = "普通住宅"
    sht_1.range("c7:c32").options(transpose=True).value = [_careful_1 for i in range(25)]

    wb_2 = app.books.open("../Excel_Test/10月高新区住宅数据.xlsx")
    sht_2 = wb_2.sheets['普通住宅']
    dt_data = sht_2.range("c2:c653").value
    #项目名称
    gx_peoject = []
    for i in dt_data:
        if i not in gx_peoject:
            gx_peoject.append(i)


    #获取套数
    count_1 = []
    for i in gx_peoject:
        variable = dt_data.count(i)
        count_1.append(variable)


    #面积
    data_1 = pd.read_excel('../Excel_Test/10月高新区住宅数据.xlsx', sheet_name="普通住宅")
    all_area = []
    for i in gx_peoject:
        some = data_1[(data_1['项目名称'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['建筑面积(平米)'].sum()
        all_area.append(Sum)



    #最高价
    up_price = []
    for i in gx_peoject:
        some = data_1[(data_1['项目名称'] == i)]

        # 指定套数列求和
        Sum = some['单价'].max()
        up_price.append(Sum)

    #最低价
    down_price = []
    for i in gx_peoject:
        some = data_1[(data_1['项目名称'] == i)]

        # 指定套数列求和
        Sum = some['单价'].min()
        down_price.append(Sum)

    #均价
    # 获取平均值
    average_1 = []
    for i in gx_peoject:
        some = data_1[(data_1['项目名称'] == i)]
        # 均价
        aver_1 = some['单价']
        area_1 = some['建筑面积(平米)']
        sum_1 = some['建筑面积(平米)'].sum()
        sum_11 = 0

        for Aver, Area_1 in zip(aver_1, area_1):
            sum_11 += Aver * Area_1

        result = sum_11/sum_1
        result = round(result, 3)
        # 平均值我这里保留的三位小数可以调整round的参数进行调整
        average_1.append(result)

    pt_name_three = sht_two.range("a2:a93").value
    pt_name_three[47] = "安建金域澜庭"
    pt_name_three[23] = '安康鼎兴苑'
    pt_name_three[4] = "安康高新总部经济与科技金融聚集区住宅项目"
    pt_name_three[36] = "安康恒大未来城"
    pt_name_three[46] = "安康金力源名苑"
    pt_name_three[88] = "安康中梁宸院"
    pt_name_three[89] = "安康中梁御墅花园"
    pt_name_three[14] = "高新波尔多·智博园公园组团"
    pt_name_three[26] = "高新观澜"
    pt_name_three[40] = "金科集美郡"
    pt_name_three[22] = "开亮崇德苑"
    pt_name_three[17] = "长兴嘉苑二期"
    pt_name_three[71] = "长兴天元郡"
    pt_name_three[90] = "中元.北城中央"

    plates_1 = sht_two.range("b2:b93").value
    # print(pt_name_two,plates)
    dic_1 = {}
    for pt_names, plate in zip(pt_name_three, plates_1):
        if pt_names in gx_peoject:
            dic_1[pt_names] = plate

    two_list = []
    for i in gx_peoject:
        if i in dic_1.keys():
            two_list.append(dic_1[i])
    # print(two_list)




    sht_1.range("f7:f32").options(transpose=True).value = gx_peoject
    sht_1.range("g7:g32").options(transpose=True).value = count_1
    sht_1.range("h7:h32").options(transpose=True).value = all_area
    sht_1.range("i7:i32").options(transpose=True).value = up_price
    sht_1.range("j7:j32").options(transpose=True).value = down_price
    sht_1.range("k7:k32").options(transpose=True).value = average_1
    sht_1.range("e7:e32").options(transpose=True).value = two_list
    wb_1.save("10月高新区住宅详情表_.xlsx")
    wb_1.close()
    wb.save("10月汉滨区住宅成交数据_.xlsx")

    # 关闭excel程序
    wb.close()
    app.quit()

