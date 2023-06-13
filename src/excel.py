class Excel(object):
    """
    整个excel文档定义
    """

    def __init__(self, file_name, host):
        self.has_bg_font_size = 9
        self.host = host
        self.bg_color = '0070c0'
        self.has_bg_font_color = 'FFFFFF'
        self.has_bg_font_name = '微软雅黑'
        self.font_color = '000000'
        self.file_name = file_name

        self.resp_map = {
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


class Sheet(object):
    """
    sheet页定义
    """

    def __init__(self, parent_title, title, name, detail, url, method, req, resp):
        self.parent_title = parent_title
        self.title = title
        self.name = name
        self.detail = detail
        self.url = url
        self.method = method
        self.req = req
        self.resp = resp


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
        self.schema_value = schema_value
        self.deep = deep

    def __str__(self) -> str:
        string = """name: {}, desc: {}, type: {}, deep: {}"""
        return string.format(self.name, self.desc, self.type, self.deep)
