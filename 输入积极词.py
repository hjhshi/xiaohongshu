file_path = "positive_words.txt"

# 检查文件末尾是否存在逗号
with open(file_path, "r", encoding="utf-8") as file:
    content = file.read()
    if not content.endswith(","):
        with open(file_path, "a", encoding="utf-8") as append_file:
            append_file.write(",")  # 如果不存在逗号，则先添加逗号

# 要求用户输入词语
words = input("请输入词语，用空格隔开：")

# 处理输入的词语
words_list = words.split()
words_str = ", ".join(['"{}"'.format(word) for word in words_list])

# 将词语写入文件
with open(file_path, "a", encoding="utf-8") as file:
    file.write(words_str)

print("词语已成功写入文件。")
