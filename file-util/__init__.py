import os


# Python根据条件修改目录里的文件名:将不想要的删去或者替换掉

# 设定文件路径
def rename(path):
    # 对目录下的文件进行遍历
    for file in os.listdir(path):
        # 判断是否是文件（查找以QL开头以.rmvb结尾的文件）
        if file.endswith(".flac"):
            # 设置新文件名
            new_name = file.replace(" ", "")
            # 重命名
            os.rename(os.path.join(path, file), os.path.join(path, new_name))


if __name__ == '__main__':
    catalog = '/Users/kevin/Documents/temp'
    rename(catalog)
