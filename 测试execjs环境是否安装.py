import execjs

# 创建一个JavaScript运行时环境
runtime = execjs.get()
# 使用运行时环境执行JavaScript代码
result = runtime.eval("'Hello, world!'")

print(result)  # 输出: Hello, world!
