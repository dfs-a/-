import xlwings as xw
"""
思路：
    1，取出预售表中的项目名称与板块对应表中的项目名称
    2，将两张表中的项目名称作比较，取板块对应表中的板块
    
"""
# 实例化应用
app = xw.App(visible=False,add_book=False)
# 工作簿
wb = app.books.add()
# 工作表
sht1 = wb.sheets["sheet1"]
class presell:
    def __init__(self):
        # 打开excel表格文件   P_l ---> Presell_list
        self.P_l = app.books.open("../Excel_Test/9月预售信息_原.xlsx")

        # 打开板块对应表 Ple_li ---> plate_list
        self.Ple_li = app.books.open('../Excel_Test/板块对应表.xlsx')

    def Get_project_name(self):
        # 选中表单
        sht = self.P_l.sheets["Sheet1"]
        # 读取预售表单项目名称   pt_name --->  project_name
        pt_name = sht.range("c4:c18").value
        pt_name[1] = "滨江公馆"
        pt_name[-1] = "集美郡"
        pt_name_dict = {}
        for k,v in enumerate(pt_name):
            pt_name_dict[k] = v
        # print(pt_name_dict)
        #商业房
        business_home = sht.range("e4:e18").value
        # print(business_home)
        #住宅房
        residence_home = sht.range("f4:f18").value
        # print(residence_home)
        """完成对应表的项目名称"""
        twos = []
        ones = []
        for pt_names,busines_home,residences_home in zip(pt_name,business_home,residence_home):
            # print(pt_names,busines_home,residences_home)
            if busines_home is not None and residences_home is not None:
                twos.append(pt_names)
            else:
                ones.append(pt_names)

        li = twos*2
        for i in li:
            ones.append(i)
        ones.sort()
        return ones,pt_name,sht

        # """ ones已排好序的项目名称 """
        #                       --->此次项目扩展小知识
        # ones.sort()           --->列表排序sort()   此方法将列表排序之后是不能重新赋值给变量的，
        #  lis = sorted(ones)   --->要想把列表排序了   还把排好序后的列表赋值给变量就需要使用sorted()排序方法



    def Plates(self,ones):
        # 选中表单
        sht_two = self.Ple_li.sheets['Sheet1']
        # 读取板块表的项目名称以及板块  pt_name_two ---> project_name  plates ----> plate
        pt_name_two = sht_two.range("a2:a93").value
        pt_name_two[75] = '吾悦广场商住项目4号、5号、6号地块'
        pt_name_two[1] = '阳光尚都'
        plates = sht_two.range("b2:b93").value
        # print(pt_name_two,plates)
        dic = {}
        for pt_names,plate in zip(pt_name_two,plates):
            if pt_names in ones:
                dic[pt_names] = plate

        # print(dic)
        one_list = []
        for one in ones:
            if one in dic.keys():
                one_list.append(dic[one])
        return one_list
        # """ 完成对应表的板块片区 """

    """ 月份 """
    def Month(self):
        Month_ = "2020/9/1"
        return Month_


    def Prepare(self,pt_name,ones,sht):
        # 此函数实现开发商,楼栋号，预售套房等信息
        """ 获取预售证号 pt_name ---> 统计表的项目名称"""
        prepare_name = sht.range("k4:k18").value
        # print(prepare_name)
        prepare_dict = {}

        # pt_names[6] = '吾悦广场商住项目4号、5号、6号地块_1'
        pt_name_1 = pt_name
        pt_name_1[5] = "吾悦广场商住项目4号、5号、6号地块_1"
        pt_name_1[7] = "安康万达ONE_1"
        pt_name_1[10] = "长兴锦源_1"
        # print(pt_name_1)
        for prt_name,prepare_names in zip(pt_name_1,prepare_name):
            prepare_dict[prt_name]=prepare_names
        # print(prepare_dict)
        # print(ones)
        ones_1 = ones
        ones_1[9] = "安康万达ONE_1"
        ones_1[14] = "长兴锦源_1"
        ones_1[4] = '吾悦广场商住项目4号、5号、6号地块_1'
        ones_1[6] = '吾悦广场商住项目4号、5号、6号地块_1'
        # print(ones_1)
        prepare_key = prepare_dict.keys()

        prepare_id = []
        for i in ones_1:
            if i in prepare_key:
                prepare_id.append(prepare_dict[i])

        #楼幢号
        Floor_Building = sht.range("d4:d18").value
        Floor_dict = {}
        for prt_name,Floor_id in zip(pt_name_1,Floor_Building):
            Floor_dict[prt_name] = Floor_id
        Floor_key = Floor_dict.keys()
        Floor_id = []
        for i in ones_1:
            if i in Floor_key:
                Floor_id.append(Floor_dict[i])


        #物业类型，物业细分
        # 商业房
        business_home = sht.range("e4:e18").value
        # 住宅房
        residence_home = sht.range("f4:f18").value
        bus_id_dict = {}
        res_id_dict = {}
        for prt_name,business_homes,residences_home in zip(pt_name,business_home,residence_home):
            bus_id_dict[prt_name] = business_homes
            res_id_dict[prt_name] = residences_home


        presell_all_number = bus_id_dict.values()
        residences_all_number_1 = res_id_dict.values()




        set_1 = []
        for i in ones_1:
            if i not in set_1:
                set_1.append(i)

        #商用
        bus_dict = {}
        for i in set_1:
            if i in bus_id_dict.keys():
                bus_dict[i] = bus_id_dict[i]


        #住宅
        res_dict = {}
        for i in set_1:
            if i in res_id_dict.keys():
                res_dict[i] = res_id_dict[i]


        #商用
        dus_id = bus_dict.values()
        #住宅
        res_id = res_dict.values()
        presell_all_number_2 = []
        for dus,res in zip(dus_id,res_id):
            if dus is None and res is not None:
                presell_all_number_2.append("住宅")
            elif dus is not None and res is None:
                presell_all_number_2.append("商业")
            elif dus is not None and res is not None:
                presell_all_number_2.extend(["商业","住宅"])


        # 预售套数
        presell_number = []

        for dus_s,res_s in zip(bus_dict.values(),res_dict.values()):
            # print(dus_s,res_s)
            if dus_s is None and res_s is not None:
                presell_number.append(res_s)
            elif dus_s is not None and res_s is None:
                presell_number.append(dus_s)
            elif dus_s is not None and res_s is not None:
                presell_number.extend([dus_s,res_s])
        # print(presell_number)

        # 预售面积
        # 商业面积
        business_area = sht.range("h4:h18").value
        # 住宅面积
        residence_area = sht.range("i4:i18").value
        buss_id_dict = {}
        ress_id_dict = {}
        for prt_name, business_one_area, residences_one_area in zip(pt_name, business_area, residence_area):
            buss_id_dict[prt_name] = business_one_area
            ress_id_dict[prt_name] = residences_one_area

        #商业
        bus_all_dict = {}
        for i in set_1:
            if i in buss_id_dict.keys():
                bus_all_dict[i] = buss_id_dict[i]


        # 住宅
        res_all_dict = {}
        for i in set_1:
            if i in ress_id_dict.keys():
                res_all_dict[i] = ress_id_dict[i]
        area = []
        for dus_all,res_all in zip(bus_all_dict.values(),res_all_dict.values()):
            if dus_all is not None and res_all is None:
                area.append(dus_all)
            elif dus_all is None and res_all is not None:
                area.append(res_all)
            else:
                area.extend([dus_all,res_all])



        # 开发商
        developers = sht.range("b4:b18").value

        developers_dict = {}
        for prt_name_s,deve in zip(pt_name_1,developers):
            developers_dict[prt_name_s] = deve
        deve_keys = developers_dict.keys()
        # print(developers_dict)
        deve_list = []
        for i in ones_1:
            if i in deve_keys:
                deve_list.append(developers_dict[i])


        return prepare_id,Floor_id,presell_all_number_2,presell_number,area,deve_list




    def run(self):
        ones, pt_name, sht = self.Get_project_name()

        one_list = self.Plates(ones)
        Month_ = self.Month()
        self.Prepare(pt_name,ones,sht)
        prepare_id,Floor_id,presell_all_number_2,presell_number,area,deve_list = self.Prepare(pt_name,ones,sht)
        # print(Floor_id)
        sht1.range("a1").value = ['项目名称','月度','板块','证号','预售楼栋','物业类别','物业细分','预售套数','预售面积','开发企业']
        sht1.range("a2").options(transpose=True).value = ones
        sht1.range("c2").options(transpose=True).value = one_list
        sht1.range("b2").options(transpose=True).value = [Month_ for i in range(20)]
        sht1.range("d2").options(transpose=True).value = prepare_id
        sht1.range("e2").options(transpose=True).value = Floor_id
        sht1.range("f2").options(transpose=True).value = presell_all_number_2
        sht1.range("g2").options(transpose=True).value = presell_all_number_2
        sht1.range("h2").options(transpose=True).value = presell_number
        sht1.range("i2").options(transpose=True).value = area
        sht1.range("j2").options(transpose=True).value = deve_list
        wb.save("预售许可统计表.xlsx")


        wb.close()
        app.quit()


if __name__ == '__main__':
    prese = presell()
    prese.run()