import pandas as pd
import xlwings as xw

app = xw.App(add_book=False,visible=False)
Ple_li = app.books.open('../Excel_Test/板块对应表.xlsx')
Pre_jet = app.books.open("../Excel_Test/10月高新区汉滨区商业.xlsx")
List_data =  Pre_jet.sheets['高新区商业']
project_Name = List_data.range("c2:c73").value
Gx_bus_data = pd.read_excel('../Excel_Test/10月高新区汉滨区商业.xlsx', sheet_name="高新区商业")

def project():
    Project_Name_one = []
    for i in project_Name:
        if i not in Project_Name_one:
            Project_Name_one.append(i)



# 套数

    count_2 = []
    for i in Project_Name_one:
        Count = project_Name.count(i)
        count_2.append(Count)


# 面积

    all_area_1 = []
    for i in Project_Name_one:
        some = Gx_bus_data[(Gx_bus_data['项目名称'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['建筑面积(平米)'].sum()
        all_area_1.append(Sum)


# 最高价

    Up_Price = []
    for i in Project_Name_one:
        some = Gx_bus_data[(Gx_bus_data['项目名称'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['单价'].max()
        Up_Price.append(Sum)


# 最低价

    Down_Price = []
    for i in Project_Name_one:
        some = Gx_bus_data[(Gx_bus_data['项目名称'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['单价'].min()
        Down_Price.append(Sum)


# 均价

    average_1 = []
    for i in Project_Name_one:
        some = Gx_bus_data[(Gx_bus_data['项目名称'] == i)]
        # 均价
        aver_1 = some['单价']
        area_1 = some['建筑面积(平米)']
        sum_1 = some['建筑面积(平米)'].sum()
        sum_12 = 0

        for Aver, Area_1 in zip(aver_1, area_1):
            sum_12 += Aver * Area_1

        result = sum_12 / sum_1
        result = round(result, 3)
        # 平均值我这里保留的三位小数可以调整round的参数进行调整
        average_1.append(result)

    return Project_Name_one,count_2,all_area_1,Up_Price,Down_Price,average_1




"""----------------------------------------------------------------------------------------"""
List_data_1 =  Pre_jet.sheets['高新区商住']
project_Name_1 = List_data_1.range("c2:c98").value
Gx_bus_data_1 = pd.read_excel('../Excel_Test/10月高新区汉滨区商业.xlsx', sheet_name="高新区商住")

# 高新区商住
def GX_Business():
    Project_Name_two = []
    for i in project_Name_1:
        if i not in Project_Name_two:
            Project_Name_two.append(i)
    # print(Project_Name_two)


    # 楼房套数
    count_4 = []
    for i in Project_Name_two:
        Count = project_Name_1.count(i)
        count_4.append(Count)

    # 面积

    all_area_2 = []
    for i in Project_Name_two:
        some = Gx_bus_data_1[(Gx_bus_data_1['项目名称'] == i)]
        # print(some)Gx_bus_data_1
        # 指定套数列求和
        Sum = some['建筑面积(平米)'].sum()
        all_area_2.append(Sum)

    # 最高价

    Up_Price = []
    for i in Project_Name_two:
        some = Gx_bus_data_1[(Gx_bus_data_1['项目名称'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['单价'].max()
        Up_Price.append(Sum)


    #最低价
    Do_Price = []
    for i in Project_Name_two:
        some = Gx_bus_data_1[(Gx_bus_data_1['项目名称'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['单价'].min()
        Do_Price.append(Sum)


    # 均价

    average_2 = []
    for i in Project_Name_two:
        some = Gx_bus_data_1[(Gx_bus_data_1['项目名称'] == i)]
        # 均价
        aver_1 = some['单价']
        area_1 = some['建筑面积(平米)']
        sum_1 = some['建筑面积(平米)'].sum()
        sum_12 = 0

        for Aver, Area_1 in zip(aver_1, area_1):
            sum_12 += Aver * Area_1

        result = sum_12 / sum_1
        result = round(result, 3)
        # 平均值我这里保留的三位小数可以调整round的参数进行调整
        average_2.append(result)
    return Project_Name_two,count_4,all_area_2,Up_Price,Do_Price,average_2

"""--------------------------------------------------"""
List_data_2 =  Pre_jet.sheets['汉滨区商业']
project_Name_2 = List_data_2.range("a2:a8").value
HB_bus_data_2 = pd.read_excel('../Excel_Test/10月高新区汉滨区商业.xlsx', sheet_name="汉滨区商业")
def Pro_ject():
    HB_business = []
    for i in project_Name_2:
        if i not in HB_business:
            HB_business.append(i)

    #套数
    all_building = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['套数'].sum()
        all_building.append(Sum)

    #面积
    area_1 = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['面积'].sum()
        area_1.append(Sum)

    #最高价
    UP_price = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['最高价'].max()
        UP_price.append(Sum)

    #最低价
    DW_price = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['最高价'].min()
        DW_price.append(Sum)

    #均价
    Average = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # 均价
        aver = some['均价']
        area = some['面积']
        sum = some['面积'].sum()
        sum_1 = 0
        for Aver, Area_1 in zip(aver, area):
            sum_1 += Aver * Area_1
        result = sum_1 / sum
        result = round(result, 2)
        # 平均值我这里保留的两位小数可以调整round的参数进行调整
        Average.append(result)

    #总套数
    all_building_1 = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['总套数'].sum()
        all_building_1.append(Sum)

    #总面积
    all_area = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['总面积'].sum()
        all_area.append(Sum)

    #可售套数
    sell_all = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['可售套数'].sum()
        sell_all.append(Sum)

    #可售面积
    sell_area = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['可售面积'].sum()
        sell_area.append(Sum)

    #已售套数
    sell_already = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['已售套数'].sum()
        sell_already.append(Sum)

    #已售面积
    sell_area_1 = []
    for i in HB_business:
        some = HB_bus_data_2[(HB_bus_data_2['QubuilderName'] == i)]
        # print(some)
        # 指定套数列求和
        Sum = some['已售面积'].sum()
        sell_area_1.append(Sum)

    return HB_business,all_building,area_1,UP_price,DW_price,Average,all_building_1,all_area,sell_all,sell_area,sell_already,sell_area_1
if __name__ == '__main__':
    Project_Name_one,Count,Area,Price,D_Price,Average = project()

    sht_three = Ple_li.sheets['Sheet1']
    pt_name_three = sht_three.range("a2:a93").value
    pt_name_three[90] = "中元.北城中央"

    plates = sht_three.range("b2:b93").value
    # print(pt_name_two,plates)
    dic = {}
    for pt_names, plate in zip(pt_name_three, plates):
        if pt_names in Project_Name_one:
            dic[pt_names] = plate

    one_list = []
    for i in Project_Name_one:
        if i in dic.keys():
            one_list.append(dic[i])



    Range = len(Project_Name_one)

    wb_1 = app.books.add()
    sht_one = wb_1.sheets["sheet1"]
    sht_one.range("a1").value = ["月度", "物业类型", "物业细分", "区域", "板块", "项目名称", "套数", "面积", "最高价",
                             "最低价", "均价", "总套数", "总面积", "可销售套数", "可销售面积", "已销售套数", "已销售面积", "备注"]

    _date_time = "2020/10/1"
    sht_one.range("a2:a{}".format(Range)).options(transpose=True).value = [_date_time for i in range(Range)]
    _type = "商业"
    sht_one.range("b2:b{}".format(Range)).options(transpose=True).value = [_type for i in range(Range)]
    sht_one.range("c2:c{}".format(Range)).options(transpose=True).value = [_type for i in range(Range)]
    _site = "高新区"
    sht_one.range("d2:d{}".format(Range)).options(transpose=True).value = [_site for i in range(Range)]
    sht_one.range("f2:f{}".format(Range)).options(transpose=True).value = Project_Name_one
    sht_one.range("g2:g{}".format(Range)).options(transpose=True).value = Count
    sht_one.range("h2:h{}".format(Range)).options(transpose=True).value = Area
    sht_one.range("i2:i{}".format(Range)).options(transpose=True).value = Price
    sht_one.range("j2:j{}".format(Range)).options(transpose=True).value = D_Price
    sht_one.range("k2:k{}".format(Range)).options(transpose=True).value = Average
    sht_one.range("e2:e{}".format(Range)).options(transpose=True).value = one_list

    # 第二张表
    wb = app.books.add()
    # 工作表
    sht = wb.sheets["sheet1"]
    sht.range("a1").value = ["月度", "物业类型", "物业细分", "区域", "板块", "项目名称", "套数", "面积", "最高价",
                                 "最低价", "均价", "总套数", "总面积", "可销售套数", "可销售面积", "已销售套数", "已销售面积", "备注"]



    Pre_ject,count_4,all_area_2,Up_Price,Do_Price,average_2 = GX_Business()

    Range_1 = len(list(Pre_ject))
    _date_time = "2020/10/1"
    sht.range("a2:a{}".format(Range_1)).options(transpose=True).value = [_date_time for i in range(Range_1)]
    _type = "商业"
    sht.range("b2:b{}".format(Range_1)).options(transpose=True).value = [_type for i in range(Range_1)]
    _type_1 = "商住"
    sht.range("c2:c{}".format(Range_1)).options(transpose=True).value = [_type_1 for i in range(Range_1)]
    _site = "高新区"
    sht.range("d2:d{}".format(Range_1)).options(transpose=True).value = [_site for i in range(Range_1)]
    sht.range("f2:f{}".format(Range)).options(transpose=True).value = Pre_ject
    sht.range("g2:g{}".format(Range)).options(transpose=True).value = count_4
    sht.range("h2:h{}".format(Range)).options(transpose=True).value = all_area_2
    sht.range("i2:i{}".format(Range)).options(transpose=True).value = Up_Price
    sht.range("j2:j{}".format(Range)).options(transpose=True).value = Do_Price
    sht.range("k2:k{}".format(Range)).options(transpose=True).value = average_2




    #第三张表格



    wb1 = app.books.add()
    sht1 = wb1.sheets["sheet1"]
    sht1.range("a1").value = ["月度", "物业类型", "物业细分", "区域", "板块", "项目名称", "套数", "面积", "最高价",
                             "最低价", "均价", "总套数", "总面积", "可销售套数", "可销售面积", "已销售套数", "已销售面积", "备注"]
    HB_business, all_building, area_1, UP_price, DW_price, Average, all_building_1, all_area, sell_all, sell_area, sell_already, sell_area_1 = Pro_ject()
    print(area_1)
    Range_2 = len(HB_business)

    pt_name_three_1 = sht_three.range("a2:a93").value
    pt_name_three_1[11] = "南龙滨江公馆"
    pt_name_three_1[52] = "康泰园小区"
    pt_name_three_1[87] = "御景华府小区"
    pt_name_three_1[75] = "安康吾悦广场商住项目"

    plates_one = sht_three.range("b2:b93").value
    # print(pt_name_two,plates)
    dic = {}
    for pt_names, plate in zip(pt_name_three_1, plates_one):
        if pt_names in HB_business:
            dic[pt_names] = plate

    three_list_1 = []
    for i in HB_business:
        if i in dic.keys():
            three_list_1.append(dic[i])




    _date_time = "2020/10/1"
    sht1.range("a2:a{}".format(Range_2)).options(transpose=True).value = [_date_time for i in range(Range_2)]
    _type = "商业"
    sht1.range("b2:b{}".format(Range_2)).options(transpose=True).value = [_type for i in range(Range_2)]
    _type_1 = "商业"
    sht1.range("c2:c{}".format(Range_2)).options(transpose=True).value = [_type_1 for i in range(Range_2)]
    _site = "汉滨区"
    sht1.range("d2:d{}".format(Range_2)).options(transpose=True).value = [_site for i in range(Range_2)]
    sht1.range("f2:f{}".format(Range_2)).options(transpose=True).value = HB_business
    sht1.range("e2:e{}".format(Range_2)).options(transpose=True).value = three_list_1
    sht1.range("g2:g{}".format(Range_2)).options(transpose=True).value = all_building
    sht1.range("h2:h{}".format(Range_2)).options(transpose=True).value = area_1
    sht1.range("i2:i{}".format(Range_2)).options(transpose=True).value = UP_price
    sht1.range("j2:j{}".format(Range_2)).options(transpose=True).value = DW_price
    sht1.range("k2:k{}".format(Range_2)).options(transpose=True).value = Average
    sht1.range("l2:l{}".format(Range_2)).options(transpose=True).value = all_building_1
    sht1.range("m2:m{}".format(Range_2)).options(transpose=True).value = all_area
    sht1.range("n2:n{}".format(Range_2)).options(transpose=True).value = sell_all
    sht1.range("o2:o{}".format(Range_2)).options(transpose=True).value = sell_area
    sht1.range("p2:p{}".format(Range_2)).options(transpose=True).value = sell_already
    sht1.range("q2:q{}".format(Range_2)).options(transpose=True).value = sell_area_1
    wb_1.save("高新区商业.xlsx")
    wb.save("高新区商住.xlsx")
    wb1.save("汉滨区商业.xlsx")
    wb.close()
    wb_1.close()
    wb1.close()
    app.quit()