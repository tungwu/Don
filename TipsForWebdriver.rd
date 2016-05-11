
selenium2 webdriver for python Tips

1. Mac环境最新驱动包(Chrome) 链接: http://chromedriver.storage.googleapis.com/index.html
  将下载的最新版本的驱dong放在 /usr/bin 目录下

2.浏览器内窗口切换 switch_to_window   获取窗口句柄 browser.window_handles[1] (0 第一个窗口 1 第二窗口 ...)
  获取句柄后切换窗口一起写 browser.switch_to_window(browser.window_handles[1])

3. 对于某些页面(浮层)的隐藏元素(通过 find 获取不到元素),可以直接进行JS注入, 借助JavaScript实现隐藏元素的点击操作
   browser.execute_script("document.getElementsByClassName('xxxxxxxxx')[0].click()")
   同样此方法也适用于页面菜单的下拉选项

4. 下拉选项的操作思路是 先定位到下拉菜单 在这个基础上再定位展开的选项 最后click一次.
   browser.find_element_by_id("下拉菜单").find_element_by_xpath("下拉选项").click()

5. python的跨目录import 需要在被调用文件的所在目录创建一个__init__.py文件 (空文件也可以)