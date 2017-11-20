# WordDocument

这个是我导师交给我的一个项目，在MFC环境下自动生成一份Word文档。方法如下：
1. 创建word文档需要几个接口类，常用application，document，documents，selection等。但word的功能复杂，要认识到每一个类的功能是不可能的。常用的方法是在word的调用宏的录制功能。通过录制的VB代码可以近似找到 相应的C++类
2. 在调用word的接口程序时需使用MSWORD.OLB，其路径位置是C:\Program Files (x86)\Microsoft Office\Office12
3. 在调用类向导后应用“类型库中的MFC类”选项
4. 所有导入的word接口类都会在都文件前附有`#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace`,将其删掉并添加`#include <afxdisp.h>`