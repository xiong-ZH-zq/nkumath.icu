"""
利用选课调查表格生产 markdown 文件的  python 脚本.

需要的选课调查表格表头为：
course, 评价1, 评价2, ... , 评价8
"""
import pandas as pd

# 导入选课表格
df = pd.read_excel(r'data\数统选课指南.xlsx')

for index, row in df.iterrows():
    course_id, course_name, course_giver = row['course'].split(' ',2)
    filename = f"docs\数院选课指南\{course_id}-{course_name}.md"  

    print(course_name)
    
    markdown_content = f"""---
comments: true
---

>[!info]+ 课程基本信息
>
> - 课程名称：{course_name}
> - 授课老师：{course_giver}
\n"""
    with open(filename, 'w') as f:
        for i in range(1,8):
           markdown_content += f"## 评价{i}\n\n{row[f'评价{i}']}\n"  
        f.write(markdown_content)
  
print("Markdown 文件已生成！")