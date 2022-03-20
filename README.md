# PowerPoint对齐工具增强版（for Windows）

![overview](./Resources/shotcut.jpg)

[插件下载链接](https://github.com/Juzhan/AlignTool/releases/tag/on_publish)

## 简介
<!-- ![detial](./Resources/shotcut_detial.jpg) -->

由于原版的对齐无法像Ai那样限定对齐对象，所以有了这个小插件。

实现了类似Adobe illusrator的对齐方式，选择多个形状后，它们会以最后被选中的形状作为基准进行对齐。


除此了增强原本对齐方式外，还增加了以将整个集合作为整体与页面对齐的功能，不用将多个形状组合后才能进行对齐。

也实现了一个在原始位置上复制粘贴形状的功能。

![detial](./Resources/show.gif)

## 安装方式

* 使用VSTO创建的小插件，点击 setup.exe 安装。安装插件后会出现在“开始”工具栏的最后，可以右键添加到快速使用栏。

## 卸载方式

1. 打开ppt,进入 文件 > 选项 > 加载项，这里可以看到安装了AlignTool，选中该插件，下方有个“转到”按钮。
2. 然后会跳出 COM 加载项，选择 AlignTool，点击右边的“删除”按钮去除该插件。
3. 进入 控制面板 > 程序 > 程序和功能，找到 AlignTool 软件，点击卸载。