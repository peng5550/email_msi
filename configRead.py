import configparser


class ReadConfig:
    """定义一个读取配置文件的类"""

    def __init__(self, configpath=None):
        if not configpath:
            return

        self.cf = configparser.ConfigParser()
        self.cf.read(configpath)

    def get_cfg(self, name, param):
        value = self.cf.get(name, param)
        return value



if __name__ == '__main__':
    res = ReadConfig("./config.ini")
    ress = res.get_cfg("template", "file_name")
    print(ress)