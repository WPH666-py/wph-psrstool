import os
import sys
import shutil
import zipfile
import rarfile
from pathlib import Path

class 水哥工具箱:
    def __init__(self):
        self.current_mode = "单文件"  # 默认模式：单文件转换
        self.create_directories()
    
    def create_directories(self):
        """创建必要的目录结构"""
        directories = ["文件上传区", "文件输出区"]
        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"已创建目录: {directory}")
    
    def show_main_menu(self):
        """显示主菜单"""
        print("\n" + "="*70)
        print("水哥工具箱 - 多功能文件转换工具")
        print('水哥工具箱，简称水箱，水柜，开个玩笑！')
        print("="*70)
        print(f"当前模式: {self.current_mode}转换")
        print("-"*70)
        if self.current_mode == "单文件":
            self.show_single_file_menu()
        else:
            self.show_folder_menu()
        
        print("-"*70)
        print("系统功能:（输入数字或指令转换）")
        print("-1. 返回上一级")
        print("-2 或 sw-fdfl - 切换模式 (当前: {})".format(self.current_mode))
        print("9. 显示作者信息")
        print("0. 退出程序")
        print("="*70)
    
    def show_single_file_menu(self):
        """显示单文件转换菜单"""
        print("单文件转换功能:（输入数字转换）")
        print("1. Word 转成 WPS")
        print("2. WPS 转成 Word")
        print("3. Excel 转成 CSV")
        print("4. CSV 转成 Excel")
    
    def show_folder_menu(self):
        """显示文件夹转换菜单"""
        print("文件夹转换功能:（输入数字转换）")
        print("5. RAR 解压成文件夹")
        print("6. 文件夹打包成 RAR")
        print("7. ZIP 解压成文件夹")
        print("8. 文件夹打包成 ZIP")
        print("-"*70)
        print("查询和删除功能:（输入指令即可）")
        print("up-fl    - 查询上传区所有文件")
        print("dn-fl    - 查询输出区所有文件")
        print("del-up   - 删除上传区所有文件")
        print("del-dn   - 删除输出区所有文件")
    
    def get_user_choice(self):
        """获取用户选择"""
        valid_choices = ['-1', '-2', '0', '9', 'sw-fdfl']
        
        if self.current_mode == "单文件":
            valid_choices.extend(['1', '2', '3', '4'])
        else:
            valid_choices.extend(['5', '6', '7', '8', 'up-fl', 'dn-fl', 'del-up', 'del-dn'])
        
        while True:
            choice = input("请输入选择: ").strip().lower()
            if choice in valid_choices:
                return choice
            else:
                print("输入无效，请输入有效的选项")
    
    def get_file_name(self):
        """获取文件名"""
        while True:
            filename = input("请输入文件名 (包含扩展名): ").strip()
            if filename:
                return filename
            else:
                print("文件名不能为空")
    
    def get_file_path(self):
        """获取文件或文件夹的相对路径"""
        while True:
            path = input("请输入相对路径: ").strip()
            if path:
                return path
            else:
                print("路径不能为空")
    
    def check_file_exists(self, file_path, source_dir="文件上传区"):
        """检查文件或文件夹是否存在"""
        full_path = os.path.join(source_dir, file_path)
        if os.path.exists(full_path):
            return full_path
        else:
            print(f"路径 '{file_path}' 在 '{source_dir}' 目录中不存在")
            return None
    
    # 单文件转换功能
    def word_to_wps_conversion(self, input_file, output_dir="文件输出区"):
        """Word 转 WPS 转换逻辑"""
        filename = os.path.basename(input_file)
        name_without_ext = os.path.splitext(filename)[0]
        output_file = os.path.join(output_dir, f"{name_without_ext}.wps")
        
        try:
            shutil.copy2(input_file, output_file)
            print(f"Word 文件已转换为 WPS 格式: {output_file}")
            return True
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    def wps_to_word_conversion(self, input_file, output_dir="文件输出区"):
        """WPS 转 Word 转换逻辑"""
        filename = os.path.basename(input_file)
        name_without_ext = os.path.splitext(filename)[0]
        output_file = os.path.join(output_dir, f"{name_without_ext}.docx")
        
        try:
            shutil.copy2(input_file, output_file)
            print(f"WPS 文件已转换为 Word 格式: {output_file}")
            return True
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    def excel_to_csv_conversion(self, input_file, output_dir="文件输出区"):
        """Excel 转 CSV 转换逻辑"""
        filename = os.path.basename(input_file)
        name_without_ext = os.path.splitext(filename)[0]
        output_file = os.path.join(output_dir, f"{name_without_ext}.csv")
        
        try:
            import pandas as pd
            
            if input_file.endswith('.xlsx') or input_file.endswith('.xls'):
                df = pd.read_excel(input_file)
                df.to_csv(output_file, index=False, encoding='utf-8-sig')
                print(f"Excel 文件已转换为 CSV 格式: {output_file}")
                return True
            else:
                print("不支持的文件格式，请提供 .xlsx 或 .xls 文件")
                return False
        except ImportError:
            print("未安装 pandas 库，无法进行 Excel 到 CSV 的转换")
            print("请运行: pip install pandas openpyxl")
            return False
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    def csv_to_excel_conversion(self, input_file, output_dir="文件输出区"):
        """CSV 转 Excel 转换逻辑"""
        filename = os.path.basename(input_file)
        name_without_ext = os.path.splitext(filename)[0]
        output_file = os.path.join(output_dir, f"{name_without_ext}.xlsx")
        
        try:
            import pandas as pd
            
            if input_file.endswith('.csv'):
                df = pd.read_csv(input_file)
                df.to_excel(output_file, index=False)
                print(f"CSV 文件已转换为 Excel 格式: {output_file}")
                return True
            else:
                print("不支持的文件格式，请提供 .csv 文件")
                return False
        except ImportError:
            print("未安装 pandas 库，无法进行 CSV 到 Excel 的转换")
            print("请运行: pip install pandas openpyxl")
            return False
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    # 文件夹转换功能
    def rar_extract(self, input_file, output_dir="文件输出区"):
        """RAR 解压成文件夹"""
        try:
            import rarfile
            
            filename = os.path.basename(input_file)
            name_without_ext = os.path.splitext(filename)[0]
            output_folder = os.path.join(output_dir, name_without_ext)
            
            with rarfile.RarFile(input_file) as rf:
                rf.extractall(output_folder)
            
            print(f"RAR 文件已解压到: {output_folder}")
            return True
        except ImportError:
            print("未安装 rarfile 库，无法进行 RAR 解压")
            print("请运行: pip install rarfile")
            return False
        except Exception as e:
            print(f"RAR 解压失败: {e}")
            return False
    
    def folder_to_rar(self, input_folder, output_dir="文件输出区"):
        """文件夹打包成 RAR"""
        try:
            import rarfile
            
            folder_name = os.path.basename(input_folder)
            output_file = os.path.join(output_dir, f"{folder_name}.rar")
            
            with rarfile.RarFile(output_file, 'w') as rf:
                for root, dirs, files in os.walk(input_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, input_folder)
                        rf.write(file_path, arcname)
            
            print(f"文件夹已打包为 RAR: {output_file}")
            return True
        except ImportError:
            print("未安装 rarfile 库，无法进行 RAR 打包")
            print("请运行: pip install rarfile")
            return False
        except Exception as e:
            print(f"RAR 打包失败: {e}")
            return False
    
    def zip_extract(self, input_file, output_dir="文件输出区"):
        """ZIP 解压成文件夹"""
        try:
            filename = os.path.basename(input_file)
            name_without_ext = os.path.splitext(filename)[0]
            output_folder = os.path.join(output_dir, name_without_ext)
            
            with zipfile.ZipFile(input_file, 'r') as zf:
                zf.extractall(output_folder)
            
            print(f"ZIP 文件已解压到: {output_folder}")
            return True
        except Exception as e:
            print(f"ZIP 解压失败: {e}")
            return False
    
    def folder_to_zip(self, input_folder, output_dir="文件输出区"):
        """文件夹打包成 ZIP"""
        try:
            folder_name = os.path.basename(input_folder)
            output_file = os.path.join(output_dir, f"{folder_name}.zip")
            
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(input_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, input_folder)
                        zf.write(file_path, arcname)
            
            print(f"文件夹已打包为 ZIP: {output_file}")
            return True
        except Exception as e:
            print(f"ZIP 打包失败: {e}")
            return False
    
    # 查询和删除功能
    def list_files_in_directory(self, directory):
        """列出目录中的所有文件和文件夹"""
        if not os.path.exists(directory):
            print(f"目录 '{directory}' 不存在")
            return []
        
        items = []
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            if os.path.isfile(item_path):
                items.append((item, "文件"))
            else:
                items.append((item, "文件夹"))
        
        return items
    
    def show_files_in_upload(self):
        """显示上传区所有文件"""
        print("\n文件上传区内容:")
        print("-"*40)
        items = self.list_files_in_directory("文件上传区")
        
        if not items:
            print("上传区为空")
        else:
            for i, (name, file_type) in enumerate(items, 1):
                print(f"{i}. {name} ({file_type})")
        print("-"*40)
    
    def show_files_in_download(self):
        """显示输出区所有文件"""
        print("\n文件输出区内容:")
        print("-"*40)
        items = self.list_files_in_directory("文件输出区")
        
        if not items:
            print("输出区为空")
        else:
            for i, (name, file_type) in enumerate(items, 1):
                print(f"{i}. {name} ({file_type})")
        print("-"*40)
    
    def delete_upload_files(self):
        """删除上传区所有文件"""
        upload_dir = "文件上传区"
        if not os.path.exists(upload_dir):
            print("上传区目录不存在")
            return False
        
        try:
            for item in os.listdir(upload_dir):
                item_path = os.path.join(upload_dir, item)
                if os.path.isfile(item_path):
                    os.remove(item_path)
                else:
                    shutil.rmtree(item_path)
            print("上传区所有文件和文件夹已删除")
            return True
        except Exception as e:
            print(f"删除失败: {e}")
            return False
    
    def delete_download_files(self):
        """删除输出区所有文件"""
        download_dir = "文件输出区"
        if not os.path.exists(download_dir):
            print("输出区目录不存在")
            return False
        
        try:
            for item in os.listdir(download_dir):
                item_path = os.path.join(download_dir, item)
                if os.path.isfile(item_path):
                    os.remove(item_path)
                else:
                    shutil.rmtree(item_path)
            print("输出区所有文件和文件夹已删除")
            return True
        except Exception as e:
            print(f"删除失败: {e}")
            return False
    
    def switch_mode(self):
        """切换模式"""
        if self.current_mode == "单文件":
            self.current_mode = "文件夹"
            print("已切换到文件夹转换模式")
        else:
            self.current_mode = "单文件"
            print("已切换到单文件转换模式")
    
    def show_author_info(self):
        """显示作者信息"""
        print("\n" + "*"*60)
        print("开发者信息")
        print("*"*60)
        print("开发者：水哥")
        print("QQ：943050454@qq.com")
        print("身份：你们的学长")
        print("-"*60)
        print("感谢使用水哥工具箱！")
        print("如有问题请联系，祝您生活愉快！")
        print("*"*60)
    
    def process_single_file_conversion(self, choice, filename):
        """处理单文件转换"""
        input_file = self.check_file_exists(filename)
        if not input_file:
            return False
        
        print(f"正在处理文件: {filename}")
        
        if choice == '1':
            return self.word_to_wps_conversion(input_file)
        elif choice == '2':
            return self.wps_to_word_conversion(input_file)
        elif choice == '3':
            return self.excel_to_csv_conversion(input_file)
        elif choice == '4':
            return self.csv_to_excel_conversion(input_file)
        else:
            return False
    
    def process_folder_conversion(self, choice, file_path):
        """处理文件夹转换"""
        input_path = self.check_file_exists(file_path)
        if not input_path:
            return False
        
        print(f"正在处理: {file_path}")
        
        if choice == '5':
            return self.rar_extract(input_path)
        elif choice == '6':
            return self.folder_to_rar(input_path)
        elif choice == '7':
            return self.zip_extract(input_path)
        elif choice == '8':
            return self.folder_to_zip(input_path)
        else:
            return False
    
    def process_query_command(self, choice):
        """处理查询命令"""
        if choice == 'up-fl':
            self.show_files_in_upload()
            return True
        elif choice == 'dn-fl':
            self.show_files_in_download()
            return True
        elif choice == 'del-up':
            return self.delete_upload_files()
        elif choice == 'del-dn':
            return self.delete_download_files()
        else:
            return False
    
    def main(self):
        """主程序"""
        print("水哥工具箱启动...")
        
        while True:
            self.show_main_menu()
            choice = self.get_user_choice()
            
            if choice == '0':
                print("感谢使用水哥工具箱，再见！")
                break
            elif choice == '-1':
                print("返回上一级")
                continue
            elif choice in ['-2', 'sw-fdfl']:
                self.switch_mode()
                continue
            elif choice == '9':
                self.show_author_info()
                continue
            
            # 处理单文件转换
            if self.current_mode == "单文件" and choice in ['1', '2', '3', '4']:
                print(f"\n您选择了模式 {choice}")
                print("请将文件放入 '文件上传区' 目录")
                print("然后输入文件名进行转换")
                
                filename = self.get_file_name()
                
                if self.process_single_file_conversion(choice, filename):
                    print("转换成功！")
                else:
                    print("转换失败，请检查文件是否存在或格式是否正确")
            
            # 处理文件夹转换和查询命令
            elif self.current_mode == "文件夹":
                if choice in ['up-fl', 'dn-fl', 'del-up', 'del-dn']:
                    if self.process_query_command(choice):
                        print("操作完成！")
                    else:
                        print("操作失败")
                elif choice in ['5', '6', '7', '8']:
                    print(f"\n您选择了模式 {choice}")
                    print("请将文件放入 '文件上传区' 目录")
                    print("然后输入相对路径进行转换")
                    
                    file_path = self.get_file_path()
                    
                    if self.process_folder_conversion(choice, file_path):
                        print("转换成功！")
                    else:
                        print("转换失败，请检查文件是否存在或格式是否正确")
            
            # 询问是否继续
            while True:
                continue_choice = input("\n是否继续其他操作？(y/n): ").strip().lower()
                if continue_choice in ['y', 'yes', '是']:
                    break
                elif continue_choice in ['n', 'no', '否']:
                    print("感谢使用水哥工具箱，再见！")
                    return
                else:
                    print("请输入 y(是) 或 n(否)")

def main():
    """程序入口"""
    try:
        toolbox = 水哥工具箱()
        toolbox.main()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"程序运行出错: {e}")

if __name__ == "__main__":
    main()