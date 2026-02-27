import os



class Directory(object):
    """
    通过读取project类来创建项目空白本目录结构
    """

    def __init__(self, project) -> None:
        self.project = project
        self.name = project.name        
        self.root = os.path.abspath('')
        
        self.goods = []
        seq = list(self.project.commodities.keys())
        seq.sort()
        for i in seq:
            index_now = self.project.commodities[i][-1]
            itemname_now = '-'.join(self.project.commodities[i][0].split('\n'))
            self.goods.append('.'.join([index_now, itemname_now]))
        

        

    def make_dir(self):
        """
        创建项目目录结构
        """
        
        level_1 = ['空白本', '副本1', '副本2']  # 创建一级目录列表
        level_2 = ['1.商务技术文件（投标函部分）', '2.商务技术文件（技术标部分）', '3.商务技术文件（经济标部分）',
         '4.商务技术文件（商务标部分）']  # 创建二级目录列表
        
        # 根据project信息创建三级目录列表
        level_3 = ['.合同条款偏离表', '.采购需求偏离表', '.物资投标响应相关文件', '.质量保证声明', '.包装方案', '.运输相关文件',
        '.物资自检验收方案', '.物资第三方检验相关文件', '.对外实施工作主体责任落实承诺书']
        if self.project.is_tech:
            level_3.append('.技术服务承诺')
        if self.project.is_qa:
            level_3.append('.售后服务承诺')
        if self.project.is_cc:
            level_3.append('.来华培训和接待承诺')
        for i in ['.舆情应对方案', '.风险防范化解方案', '.物资中主要标的的生产企业三体系资料', '.其它说明和资料']:
            level_3.append(i)

        path_0 = '\\'.join([self.root, '投标文件-{}'.format(self.name)])  # 确定根目录路径
        path_1 = '\\'.join([path_0, level_1[0]])
        path_2 = '\\'.join([path_1, level_2[1]])
        path_3 = '\\'.join([path_2, '3.物资投标响应相关文件'])

        for dirnow in level_1:  # 创建根目录及一级目录
            os.makedirs('\\'.join([path_0, dirnow]))

        for dirnow in level_2:  # 创建二级目录
            os.mkdir('\\'.join([path_1, dirnow]))

        for i in range(len(level_3)): # 创建三级目录
            dirnow = ''.join([str(i + 1), level_3[i]])
            pathnow = '\\'.join([path_2, dirnow])
            os.mkdir(pathnow)

        for dirnow in self.goods:  # 写入物资名文件夹
            pathnow = '\\'.join([path_3, dirnow])
            os.mkdir(pathnow)

        


        
       

