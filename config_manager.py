import json


class Setting:
    all_settings = []
    config_file_path: str = 'Resources/config.json'

    def __init__(self, root_name: str):
        assert len(root_name) > 0, f'root name is required!'
        # self.initialized = False
        self.root_name = root_name

        Setting.all_settings.append(self)

    def read_config(self):
        with open(Setting.config_file_path) as json_file:
            _json = json.load(json_file)
        json_file.close()
        if _json is None:
            print(f'No config for {self.__name__} found. Default values will be used.')
            return
        for param in self.__dict__:
            try:
                val = _json[self.root_name][param]
                self.__dict__[param] = val
            except:
                print(
                    f'{param} is not found in {self.config_file_path}. the default value of : {getattr(self, param)} used')
                continue

    def write_config(self, key, val):
        with open(Setting.config_file_path) as json_file:
            conf = json.load(json_file)
        json_file.close()
        with open(Setting.config_file_path, 'w+', encoding='utf-8') as json_file:
            try:
                conf[self.root_name][key] = val
            except:
                conf.update({self.root_name: {key: val}})
            json.dump(conf, fp=json_file, ensure_ascii=False, indent=4)
        json_file.close()


class DataSetting(Setting):
    initialized = False

    def __init__(self):
        super().__init__(root_name='data')
        self.modified_application_list_path = ''
        self.working_folder_path = ''
        self.read_config()
        self.initialized = True

    def __setattr__(self, key, value):
        if key not in self.__dict__:
            super().__setattr__(key, value)
        self.__dict__[key] = value
        if self.initialized and key != 'initialized':
            self.write_config(key, value)
