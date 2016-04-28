
selenium2 webdriver for python Tips

1. Mac环境最新驱动包(Chrome) 链接: http://chromedriver.storage.googleapis.com/index.html
  将下载的最新版本的驱dong放在 /usr/bin 目录下

2.浏览器内窗口切换 switch_to_window   获取窗口句柄 browser.window_handles[1] (0 第一个窗口 1 第二窗口 ...)
  获取句柄后切换窗口一起写 browser.switch_to_window(browser.window_handles[1])
