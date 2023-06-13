import json
import re


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
        self.interface_list = []

    def append_interface(self, item):
        self.interface_list.append(item)

    def __str__(self) -> str:
        string = """title:{}, name:{}, desc:{}, author:{}, host:{}, version:{}"""
        return string.format(self.title, self.name, self.desc, self.author, self.host, self.version)


class Interface(object):
    def __init__(self, name, desc, method, summary, url, produces, consumes):
        self.request_params = []
        self.response_params = []
        self.url = url
        self.name = name
        self.desc = desc
        self.method = method
        self.summary = summary
        self.produces = produces
        self.consumes = consumes

    def append_request(self, item):
        self.request_params.append(item)

    def append_response(self, item):
        self.response_params.append(item)


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
                exit(0)
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
            for child in tag.get('childrens'):
                desc = child.get('description').replace("<p>", "").replace("</p>", "")
                method = child.get('methodType')
                url = child.get('showUrl')
                summary = child.get('summary').replace('[', '【').replace(']', '】')
                produces = child.get('produces')
                consumes = child.get('consumes')
                req_params = child.get('reqParameters')
                resp_params = child.get('multipData').get('responseParameters')
                interface = Interface(name=one_title, desc=desc, method=method, summary=summary, url=url,
                                      produces=produces, consumes=consumes)
                interface.append_request(self.parse_child(req_params))
                interface.append_response(self.parse_child(resp_params))
                self.project.append_interface(interface)

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
