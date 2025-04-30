# Webcrawler_Name_info
批量获取专家信息。访问“https://www.letpub.com.cn/nsfcfund_search.php”（letpub）  
并从中获取名字所对应的全部专家（有重名），并保存每个专家所对应的研究方向、学科代码、学校以及帽子


---
## 函数说明：  
### get_data  
**输入:**  
name => 专家的名字  
company => 所在单位  
begin_year => 起始年份  
end_year +> 终止年份  
这些参数将被置入post所携带的data中，并完成Response的获取。  
**返回:**  
letpub上返回的信息。


### append_to_excel  
**输入：**  
name =>  专家名字  
title =>  项目名称  
codes =>  学科代码  
project_type =>  项目的等级  
company_names =>  单位名称  
file_name =>  保存到的文件名  
**返回:**  
将文件保存到excel表中  

## 技术流程：  
1.创建储存xlsx  
2.遍历原始文件  
3.对每一个名字，有：（向letpub请求数据 => 解析html文本 => 使用xpath提取关键信息 => 向excel储存信息）  
4.遍历完成，结束循环  
