# 基于GIS 的物流系统规划设计
特别说明：本产品只为沟通和学习交流使用，任何形式的商业用途都是不被允许的，对于产品可能带来的法律问题，开发者保留最终的解释权。
##项目简介
本项目基于GIS的空间分析功能，以兰州申通物流为例，利用MapObject和VB构建GIS物流配送路线规划系统。
系统大致分为登陆窗体模块,主窗体模块,空间显示模块,查询模块,空间分析模块等。
首先构建兰州市城关区道路交通网络，通过Access创建申通物流客户的姓名等基本信息和空间位置数据库。
通过Flyod最短路径算法对物流配送线路进行划分，以期优化物流配送路线，提高物流配送的效益。 
##数据库设计
1.	空间实体分别用点、线、面图层表达。
2.	使用关系数据库Access，每个空间实体添加id号和名称属性项
3．对空间数据进行分类分级,统一采用五位编码方案,前两位表示空间数据的类别,后三位表示该空间数据在所属类别中的编码.用五位的编码作为该空间数据唯一的标志符,便于对数据进行存储,处理,查询,分析.编码示例35001.
##系统设计的基本构架
![image](https://raw.githubusercontent.com/mengyingliu/LogisticSystem/master/%E7%89%A9%E6%B5%81%E9%85%8D%E9%80%81%E7%B3%BB%E7%BB%9F/%E5%9B%BE%E7%89%87%E5%9B%BE%E6%A0%87/%E7%B3%BB%E7%BB%9F%E6%9E%B6%E6%9E%84.png)
##各模块功能显示
![image](https://raw.githubusercontent.com/mengyingliu/LogisticSystem/master/%E7%89%A9%E6%B5%81%E9%85%8D%E9%80%81%E7%B3%BB%E7%BB%9F/%E5%9B%BE%E7%89%87%E5%9B%BE%E6%A0%87/%E5%90%84%E6%A8%A1%E5%9D%97%E5%8A%9F%E8%83%BD%E6%98%BE%E7%A4%BA.png)
##最短路径查询界面
![image](https://raw.githubusercontent.com/mengyingliu/LogisticSystem/master/%E7%89%A9%E6%B5%81%E9%85%8D%E9%80%81%E7%B3%BB%E7%BB%9F/%E5%9B%BE%E7%89%87%E5%9B%BE%E6%A0%87/%E7%94%A8%E6%88%B7%E4%BF%A1%E6%81%AF%E6%9F%A5%E8%AF%A2.png)
##地图索引
![image](https://raw.githubusercontent.com/mengyingliu/LogisticSystem/master/%E7%89%A9%E6%B5%81%E9%85%8D%E9%80%81%E7%B3%BB%E7%BB%9F/%E5%9B%BE%E7%89%87%E5%9B%BE%E6%A0%87/%E5%9C%B0%E5%9B%BE%E7%B4%A2%E5%BC%95.png)
