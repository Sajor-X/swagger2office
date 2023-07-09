import json
import re
from typing import List


class Project(object):
    """
    项目
    """

    def __init__(self, title, name, desc, author, host, version):
        """
        :param title: 标题
        :param name: 分组名称
        :param desc: 简介
        :param author: 作者
        :param host: 地址
        :param version: 版本
        """
        self.title = title
        self.name = name
        self.desc = desc
        self.author = author
        self.host = host
        self.version = version
        self.interface_group_list: List[Group] = []
        self.separator = "-"

        self.resp_map = {
            'pageNum': {
                'chinese': '分页页数',
                'require': True,
                'desc': '分页参数中的页数，默认为第一页'
            },
            'total': {
                'chinese': '总数',
                'require': True,
                'desc': '总条数，可搭配分页使用'
            },
            'list': {
                'chinese': '具体数据',
                'require': True,
                'desc': '具体数据列表'
            },
            'pageSize': {
                'chinese': '分页大小',
                'require': True,
                'desc': '分页参数中的页大小，默认为10条'
            },
            'code': {
                'chinese': '状态码',
                'require': True,
                'desc': '1成功，0失败，-1异常'
            },
            'message': {
                'chinese': '状态说明',
                'require': False,
                'desc': '异常时返回错误信息'
            },
            'data': {
                'chinese': '返回数据',
                'require': True,
                'desc': ''
            }
        }

    def append_group(self, item):
        self.interface_group_list.append(item)

    def __str__(self) -> str:
        string = """title:{}, name:{}, desc:{}, author:{}, host:{}, version:{}"""
        return string.format(self.title, self.name, self.desc, self.author, self.host, self.version)

class Group(object):
    def __init__(self, title, desc):
        self.title = title
        self.desc = desc
        self.interface_list: List[Interface] = []
    def append_interface(self, item):
        self.interface_list.append(item)


class Interface(object):
    def __init__(self, name, desc, method, group, url, produces, consumes):
        self.request_params: List[BodyItem] = []
        self.response_params: List[BodyItem] = []
        self.url = url
        self.name = name
        self.desc = desc
        self.method = method
        self.group = group
        self.produces = produces
        self.consumes = consumes

class BodyItem(object):
    """
    请求&返回体 分项
    """

    def __init__(self, child, desc, require, name, type, mode, schema_value, deep):
        self.name = name
        self.require = require
        self.desc = desc
        self.type = type
        self.child = child
        self.mode = mode
        # 判断是否为对象
        self.schema_value = schema_value
        self.deep = deep

    def __str__(self) -> str:
        string = """name: {}, desc: {}, type: {}, deep: {}"""
        return string.format(self.name, self.desc, self.type, self.deep)


class ParseAPIDoc(object):
    def __init__(self):
        self.project = None

    def file_to_json(self, file_name):
        """
        解析html文件获取接口数据
        :param file_name:
        :return:
        """
        with open(file_name, 'r', encoding='utf-8') as f:
            findall = re.findall("var datas=(.*);\n", f.read())
            if len(findall) != 1:
                raise Exception
            return json.loads(findall[0])

    def parse_instance(self, instance):
        """
        创建项目文档
        """
        title = instance.get('title')
        description = instance.get('description')
        contact = instance.get('contact')
        version = instance.get('version')
        host = instance.get('host')
        basePath = instance.get('basePath')
        name = instance.get('name')
        self.project = Project(title=title, name=name, desc=description, author=contact, host=host, version=version)

    def parse_tags(self, tags):
        """
        解析标签
        """
        pass

    def parse(self, json_data):
        """
        解析
        """
        # 解析项目文档
        instance = json_data.get('instance')
        self.parse_instance(instance=instance)

        tags = json_data.get('tags')
        self.parse_params(tags)

        print(self.project)

    def parse_child(self, datas, deep=0):
        if not datas:
            return []
        d = []
        for data in datas:
            children = self.parse_child(data.get('children'), deep + 1)
            desc = data.get('description')
            require = data.get('require')
            name = data.get('name')
            mode = data.get('in')
            type = data.get('type')
            schema_value = data.get('schemaValue')
            d.append(BodyItem(children, desc, require, name, type, mode, schema_value, deep))
            d.extend(children)
        return d

    def parse_params(self, tags):
        """
        解析参数
        :param tags:
        :return:
        """
        for tag in tags:
            one_title = tag.get('name')
            one_description = tag.get('description')
            # 创建接口组
            interface_group = Group(one_title, one_description)
            for child in tag.get('childrens'):
                desc = child.get('description').replace("<p>", "").replace("</p>", "").rstrip('\n')
                method = child.get('methodType')
                url = child.get('showUrl')
                summary = child.get('summary').replace('[', '【').replace(']', '】')
                produces = child.get('produces')
                consumes = child.get('consumes')
                req_params = child.get('reqParameters')
                resp_params = child.get('multipData').get('responseParameters')
                interface = Interface(name=summary, desc=desc, method=method, group=one_title, url=url,
                                      produces=produces, consumes=consumes)
                interface.request_params = self.parse_child(req_params)
                interface.response_params = self.parse_child(resp_params)

                interface_group.append_interface(interface)

            self.project.append_group(interface_group)

    def run(self, file_name):
        """
        解析项目对象
        """
        json_data = self.file_to_json(file_name)
        self.parse(json_data)
        return self.project


if __name__ == '__main__':
    p = ParseAPIDoc()
    project = p.run('../swagger.html')
