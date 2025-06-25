import pandas as pd
from lxml import etree
import os
import numpy as np
class XDMConverter:
    def __init__(self):
        self.var_type_map = {}
        self.var_a_map = {}

    def load_var_type_and_a_map(self, xml_file, key_words):
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(xml_file, parser)
        root = tree.getroot()

        for key in key_words:
            type_dict = {}
            lst_list = root.xpath(f'.//*[local-name() = "lst" and @name="{key}"]')
            if lst_list:
                lst = lst_list[0]
                for ctr in lst.xpath(f'./*[local-name() = "ctr"]'):
                    ctr_type = ctr.attrib.get('type')
                    type_dict['name'] = ctr_type
                    for var in ctr.xpath('./*[local-name()="var"]'):
                        name = var.attrib.get('name')
                        type_dict[name] = var.attrib.get('type')
                        a_list = []
                        for a in var.xpath('./*[local-name()="a"]'):
                            a_name = a.attrib.get('name')
                            a_type = a.attrib.get('type')
                            a_value = a.attrib.get('value', '')
                            
                            if a_name:
                                a_list.append((a_name,a_type, a_value))
                        if a_list:
                            self.var_a_map[name] = a_list
                self.var_type_map[key] = type_dict
            else:
                for ctr in root.xpath(f'.//*[local-name()="ctr" and @name="{key}"]'):
                    ctr_type = ctr.attrib.get('type')
                    type_dict['name'] = ctr_type
                    for var in ctr.xpath('./*[local-name()="var"]'):
                        name = var.attrib.get('name')
                        type_dict[name] = var.attrib.get('type')
                        a_list = []
                        for a in var.xpath('./*[local-name()="a"]'):
                            a_name = a.attrib.get('name')
                            a_value = a.attrib.get('value', '')
                            if a_name:
                                a_list.append((a_name, a_value))
                        if a_list:
                            self.var_a_map[name] = a_list
                self.var_type_map[key] = type_dict

    def xml_to_dataframe(self, xml_file, key_word):
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(xml_file, parser)
        root = tree.getroot()
        data = []

        self.var_a_map = {}  # key: var name, value: list of (a name, a value)
        lst_list = root.xpath(f'.//*[local-name() = "lst" and @name="{key_word}"]')
        if lst_list:
            type_dict = {}
            lst = lst_list[0]
            for ctr in lst.xpath(f'./*[local-name() = "ctr"]'):
                row = {'name': ctr.attrib.get('name')}
                for var in ctr.xpath('./*[local-name()="var"]'):
                    name = var.attrib.get('name')
                    value = var.attrib.get('value')
                    if value:
                        row[name] = value
                    type = var.attrib.get('type')
                    type_dict[name] = type

                    # Lưu thẻ <a> nếu có
                    a_list = []
                    for a in var.xpath('./*[local-name()="a"]'):
                        a_name = a.attrib.get('name')
                        a_value = a.attrib.get('value', '')
                        if a_name:
                            a_list.append((a_name, a_value))
                    if a_list:
                        self.var_a_map[name] = a_list

                self.var_type_map[key_word] = type_dict
                data.append(row)
        else:
            for ctr in root.xpath(f'.//*[local-name()="ctr" and @name="{key_word}"]'):
                row = {'name': ctr.attrib.get('name')}
                for var in ctr.xpath('./*[local-name()="var"]'):
                    name = var.attrib.get('name')
                    value = var.attrib.get('value')
                    if value:
                        row[name] = value
                    type_dict[name] = var.attrib.get('type')

                    # Lưu thẻ <a> nếu có
                    a_list = []
                    for a in var.xpath('./*[local-name()="a"]'):
                        a_name = a.attrib.get('name')
                        a_value = a.attrib.get('value', '')
                        if a_name:
                            a_list.append((a_name, a_value))
                    if a_list:
                        self.var_a_map[name] = a_list

                self.var_type_map[key_word] = type_dict
                data.append(row)

        return pd.DataFrame(data)



    def xml_to_excel(self, xml_file, listKey):
        if not os.path.exists(xml_file):
            raise FileNotFoundError("[ERROR] Không tìm thấy file")
        data_arr = {}
        for key in listKey:
            df = self.xml_to_dataframe(xml_file, key)
            if not df.empty:
                data_arr[key] = df

        if not data_arr:
            raise ValueError("[FAILED] Không có giá trị hợp lệ theo key_word")
        else:
            with pd.ExcelWriter(f"{os.path.splitext(xml_file)[0]}.xlsx", engine='openpyxl') as writer:
                for key, df in data_arr.items():
                    df.to_excel(writer, sheet_name=f"{key}", index=False)

    def excel_to_dataframe(self, fileExcel_name, key_words):
        data_arr = []
        for key in key_words:
            df = pd.read_excel(fileExcel_name, sheet_name=key)
            data_arr.append(df)
        return data_arr

    def excel_to_xml(self, xml_file, key_words):
        df_arr = self.excel_to_dataframe(os.path.splitext(xml_file)[0] + '.xlsx', key_words)
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(xml_file, parser)
        root = tree.getroot()

        for key, df in zip(key_words, df_arr):
            df = df.dropna(how='all')
            df.fillna(value=np.nan, inplace=True)
            df_columns = df.columns.tolist()

            # tìm node lst
            lst_list = root.xpath(f'.//*[local-name() = "lst" and @name="{key}"]')
            lst = lst_list[0] if lst_list else None

            if lst is not None:
                ctrs = lst.xpath('./*[local-name()="ctr"]')
            else:
                ctrs = root.xpath(f'.//*[local-name() = "ctr" and @name="{key}"]')

            existing_names = set()
            for ctr in ctrs:
                name = ctr.attrib.get('name')
                matching_rows = df[df['name'] == name]

                if not matching_rows.empty:
                    row = matching_rows.iloc[0]
                    existing_names.add(name)
                    for var in ctr.xpath('./*[local-name() = "var"]'):
                        var_name = var.get('name')
                        if var_name in df.columns:
                            new_val = row.get(var_name, "")
                            if pd.notna(new_val):
                                var.set('value', str(new_val).strip())
                            else:
                                var.set('value', "None")
                else:
                    parent = ctr.getparent()
                    parent.remove(ctr)

            # Thêm mới các ctr chưa có trong XML
            new_rows = df[~df['name'].isin(existing_names)]
            for _, row in new_rows.iterrows():
                new_ctr = etree.Element("ctr")
                new_ctr.set("name", str(row['name']))

                for col in df_columns:
                    if col == "name":
                        continue
                    value = row[col]
                    var_elem = etree.Element("var")
                    var_elem.set("name", col)

                    if pd.isna(value):
                        var_elem.set("value", "None")
                    elif isinstance(value, float) and value.is_integer():
                        var_elem.set("value", str(int(value)))
                    else:
                        var_elem.set("value", str(value).strip())

                    # Thêm thẻ <a> nếu có mẫu
                    if col in self.var_a_map:
                        for a_name,a_type, a_value in self.var_a_map[col]:
                            a_elem = etree.Element("a")
                            a_elem.set("name", a_name)
                            if a_type is not None:
                                a_elem.set('type',a_type)
                            if a_value is not None:
                                a_elem.set("value", str(a_value))
                            var_elem.append(a_elem)

                    new_ctr.append(var_elem)

                if lst is not None:
                    lst.append(new_ctr)
                else:
                    root.append(new_ctr)

        tree.write(xml_file, pretty_print=True, xml_declaration=True, encoding="utf-8")
        print("[SUCCESS] Ghi file XML thành công.")




    def find_node(self, root, name):
        for item in root.iter():
            tag_name = item.tag.split('}')[-1]
            if tag_name in ('lst', 'ctr') and item.get('name') == name:
                return item
        return None

    def replace_part(self, main_file: str, replace_file: str, key_word_main: str, key_word_replace: str):
        parser = etree.XMLParser(remove_blank_text=False)
        tree_main = etree.parse(main_file, parser)
        root_main = tree_main.getroot()
        tree_replace = etree.parse(replace_file, parser)
        root_replace = tree_replace.getroot()

        main_node = self.find_node(root_main, key_word_main)
        if main_node is None:
            print(f'[ERROR] Không tìm thấy từ khóa {key_word_main} trong file chính')
            return

        replace_node = self.find_node(root_replace, key_word_replace)
        if replace_node is None:
            print(f'[ERROR] Không tìm thấy từ khóa {key_word_replace} trong file thay thế')
            return

        parent = main_node.getparent()
        idx = parent.index(main_node)
        parent.remove(main_node)
        parent.insert(idx, replace_node)
        tree_main.write(main_file, pretty_print=True, xml_declaration=True, encoding='utf-8')
        print('[SUCCESS] Ghi file thành công')

   
def main():
    file = XDMConverter()
    input = file.read_input('test_input.txt')

    print(input)

    
    """
    ý tưởng xử lý phần ghi thêm file: 
    1 dataframe để lưu name và value của trong excel
    1 dict gồm:
    key: các từ khóa người dùng nhập vào
    value: 1 dict gồm các tên và thuộc tính của đối tượng chứa tên đó (giống excel nhưng thay value = type)
    đọc các thẻ <a> nằm bên trong các thẻ var, mỗi thẻ var lại có dict tương ứng gồm name và value của các thẻ a bên trong. 
    sau khi thêm node mới thì tự động thêm giống như các thẻ <a> 
    """
    # Ví dụ chạy thử:
    # file.xml_to_excel(xml_file, input['key_words'])
    # file.excel_to_xml(xml_file, input['key_words'])

if __name__ == "__main__":
    main()
