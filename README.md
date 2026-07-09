pptx_to_mermaid 工具说明文档

一款可以将 PowerPoint PPTX 中的框图、流程图形秒级转换为 Markdown 和标准 Mermaid 流程图语法的命令行工具，专门适配 AI 编程代理（Vibe Coding Agent）输入场景，让幻灯片里的业务逻辑、架构图、流程说明无需手动重绘，直接导入大模型辅助需求拆解、代码生成与设计复盘。

✨ 核心特性
精准识别PPT原生元素‌
自动提取幻灯片中的普通形状、连接线、文本层级，完美解析微软PPT原生SmartArt图形的逻辑结构，不会遗漏流程节点和关联关系。
智能方向自动检测‌
无需手动指定流程图排布，工具会自动根据所有连接线的水平/垂直跨度，智能判断生成从上到下(TD)还是从左到右(LR)的Mermaid流程图，输出排版自然流畅。
全内容无损导出‌
除了流程图之外，还会自动将幻灯片中的普通文本、嵌套列表、表格、图片引用直接转换为标准Markdown格式，完整保留PPT中的全部内容信息。
专为AI编程场景优化‌
输出格式原生兼容所有支持Mermaid渲染的AI编程助手、文档编辑器，转换后的纯文本内容没有PPT的格式冗余，大模型可以直接读取理解流程逻辑，大幅提升需求对齐效率。
轻量跨平台‌
基于C++17开发，仅依赖tinyxml2和miniz两个第三方库，支持Windows、Linux、macOS全平台编译运行，没有Python/Java等运行环境依赖。
🛠️ 编译构建说明

根据你的操作系统和编译环境，选择对应的构建命令即可：

Windows平台
MinGW / Git Bash 环境‌
执行以下命令即可生成可执行文件：
g++ -std=c++17 -O2 -Ivendor pptx_to_mermaid.cpp vendor/tinyxml2.cpp -o pptx_to_mermaid.exe -lz
MSVC 开发者命令提示符环境‌
执行以下命令编译：
cl /std:c++17 /EHsc /Ivendor pptx_to_mermaid.cpp vendor\tinyxml2.cpp
Linux平台

直接在终端执行构建命令：
g++ -std=c++17 -O2 -Ivendor pptx_to_mermaid.cpp vendor/tinyxml2.cpp -o pptx_to_mermaid -lz

📖 使用指南
基础用法

最简化的使用方式，直接输入PPTX文件路径即可，转换结果默认输出到标准控制台：
pptx_to_mermaid input.pptx

如果需要将结果保存为指定Markdown文件，直接追加输出路径即可：
pptx_to_mermaid input.pptx [output.md]

可选参数说明

支持的所有命令行选项如下：

表格
参数选项	功能说明
-s, --slides N [N ...]	仅转换指定编号的幻灯片，幻灯片序号从1开始计数，支持同时指定多个页码
-d, --direction DIR	强制指定Mermaid流程图的排布方向，可选值为TD（从上到下）、LR（从左到右）、BT（从下到上）、RL（从右到左），默认值为auto，由工具自动判断最优方向
--no-smartart	关闭SmartArt图形解析功能，跳过幻灯片中的所有SmartArt内容转换
--no-fence	输出原生Mermaid语法，不添加```mermaid代码块包裹，适合需要直接拼接Mermaid内容的场景
--one-section	默认开启，每一页幻灯片生成独立的二级Markdown标题，对应单独的内容区块
--single	将所有幻灯片的流程图合并到同一个Mermaid代码块中输出
-h, --help	查看帮助信息
常用使用示例
仅转换第1页和第3页幻灯片，输出到result.md文件：
pptx_to_mermaid deck.pptx result.md --slides 1 3
强制生成从左到右排布的流程图：
pptx_to_mermaid deck.pptx -d LR
输出无代码块包裹的原生Mermaid内容，直接传入AI编程代理：
pptx_to_mermaid design.pptx --no-fence
📋 输出规则说明

工具会自动判断每一页幻灯片的内容类型，采用不同的转换策略：

包含连接线的流程图页面‌
自动生成符合标准的Mermaid flowchart代码，所有形状会被转换为带文本标签的节点，连接线直接映射为Mermaid箭头，双向箭头会自动转换为两条互指的连线。页面中不属于流程图节点的剩余文本、表格、图片会自动追加到Mermaid代码块下方，转换为普通Markdown内容。
无流程图的普通内容页面‌
直接转换为标准Markdown格式：
幻灯片标题占位符内容自动转换为一级标题
原生PPT中的嵌套列表会根据段落缩进层级自动生成对应缩进的Markdown无序列表，保留原有层级关系
PPT中的表格自动转换为标准Markdown表格，自动转义单元格内的特殊竖线字符
幻灯片内的图片自动转换为Markdown图片引用格式
SmartArt图形‌
自动解析SmartArt的内部逻辑结构，提取所有节点文本和关联关系，生成对应的Mermaid流程图，即使部分SmartArt节点缺失文本，也会自动从形状的可视化内容中兜底提取文本。
🚀 接入Vibe Coding Agent实践指南

工具生成的输出内容可以直接输入到AI编程代理中，大幅提升开发效率：

直接将转换后的Markdown文件粘贴到AI助手对话框，大模型可以秒级理解流程图中的业务流转逻辑，无需人工逐节点描述流程。
配合--no-fence --single参数可以输出合并后的纯Mermaid流程文本，避免代码块格式干扰，让AI代理更精准地解析全量业务链路。
对于产品经理输出的PPT需求文档，全量转换后可以直接导入AI编程代理，自动生成对应业务模块的接口定义、函数逻辑和单元测试，大幅减少需求沟通成本和手动设计遗漏。
📝 注意事项
遇到特殊字符、换行符时，工具会自动对Mermaid标签进行转义，将换行转换为<br/>，避免语法错误。
连接线两端如果没有绑定到具体形状，工具会通过坐标匹配最近的形状节点，超出合理距离的悬空连线会自动丢弃，避免生成无效连线。
如果转换后提示"No box diagrams found"，说明对应PPT中没有检测到可解析的流程图元素，不会生成空输出。
