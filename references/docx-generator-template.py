"""
DSH 简历 Word (.docx) 生成脚本模板
基于 python-docx 构建符合现代商业排版规范的单页/多页简历文档
"""

import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_border(cell, **kwargs):
    """设置单元格边框：top, bottom, left, right"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = parse_xml(r'<%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' % (
                tag, nsdecls('w'), edge_data.get('val', 'single'),
                edge_data.get('sz', '4'), edge_data.get('color', 'auto')
            ))
            tcBorders.append(element)
    tcPr.append(tcBorders)

def create_resume_docx(data, output_path="resume.docx"):
    doc = Document()
    
    # 1. 页面设置：A4纸，经典窄边距（上1.5cm，下1.5cm，左1.8cm，右1.8cm）
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    
    # 2. 配色体系
    COLOR_PRIMARY = RGBColor(30, 58, 138)    # 深蓝 #1E3A8A (主标题)
    COLOR_ACCENT = RGBColor(59, 130, 246)    # 亮蓝 #3B82F6 (分割线/小标题)
    COLOR_TEXT = RGBColor(31, 41, 55)        # 深炭灰 #1F2937 (正文)
    COLOR_MUTED = RGBColor(107, 114, 128)    # 次要信息灰 #6B7280
    
    # 3. 头部信息 (Header)
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.paragraph_format.space_before = Pt(0)
    p_header.paragraph_format.space_after = Pt(4)
    
    run_name = p_header.add_run(data.get("name", "姓名"))
    run_name.font.size = Pt(20)
    run_name.font.bold = True
    run_name.font.color.rgb = COLOR_PRIMARY
    run_name.font.name = "Microsoft YaHei"
    
    # 联系方式一行居中
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(12)
    
    contact_parts = []
    if "target_role" in data: contact_parts.append(f"求职意向：{data['target_role']}")
    if "phone" in data: contact_parts.append(data["phone"])
    if "email" in data: contact_parts.append(data["email"])
    if "location" in data: contact_parts.append(data["location"])
    if "links" in data: contact_parts.append(data["links"])
    
    run_contact = p_contact.add_run("  |  ".join(contact_parts))
    run_contact.font.size = Pt(9.5)
    run_contact.font.color.rgb = COLOR_MUTED
    run_contact.font.name = "Microsoft YaHei"
    
    def add_section_header(title):
        """添加带下划线的模块标题"""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(title)
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        run.font.name = "Microsoft YaHei"
        
        # 底部下划线边框
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="12" w:space="2" w:color="3B82F6"/></w:pBdr>' % nsdecls('w'))
        pPr.append(pBdr)

    # 4. 教育背景 (Education)
    if "education" in data and data["education"]:
        add_section_header("教育背景")
        for edu in data["education"]:
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            table.columns[0].width = Inches(4.5)
            table.columns[1].width = Inches(2.37)
            
            # 左侧：学校与专业
            cell_l = table.cell(0, 0)
            p_l = cell_l.paragraphs[0]
            p_l.paragraph_format.space_before = Pt(2)
            p_l.paragraph_format.space_after = Pt(2)
            r1 = p_l.add_run(edu.get("school", ""))
            r1.font.bold = True
            r1.font.size = Pt(10)
            r1.font.color.rgb = COLOR_TEXT
            r1.font.name = "Microsoft YaHei"
            
            r2 = p_l.add_run(f"  |  {edu.get('major', '')}  |  {edu.get('degree', '')}")
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = COLOR_TEXT
            r2.font.name = "Microsoft YaHei"
            
            # 右侧：时间与地点
            cell_r = table.cell(0, 1)
            p_r = cell_r.paragraphs[0]
            p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_r.paragraph_format.space_before = Pt(2)
            p_r.paragraph_format.space_after = Pt(2)
            r_time = p_r.add_run(edu.get("date", ""))
            r_time.font.size = Pt(9.5)
            r_time.font.color.rgb = COLOR_MUTED
            r_time.font.name = "Microsoft YaHei"

    # 5. 专业技能 (Skills)
    if "skills" in data and data["skills"]:
        add_section_header("专业技能")
        for skill_cat in data["skills"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            
            run_title = p.add_run(f"• {skill_cat.get('category', '')}：")
            run_title.font.bold = True
            run_title.font.size = Pt(9.5)
            run_title.font.color.rgb = COLOR_TEXT
            run_title.font.name = "Microsoft YaHei"
            
            run_detail = p.add_run(skill_cat.get("detail", ""))
            run_detail.font.size = Pt(9.5)
            run_detail.font.color.rgb = COLOR_TEXT
            run_detail.font.name = "Microsoft YaHei"

    # 6. 工作与项目经历 (Experience / Projects)
    for section_key, section_title in [("experience", "工作经历"), ("projects", "项目经历")]:
        if section_key in data and data[section_key]:
            add_section_header(section_title)
            for item in data[section_key]:
                table = doc.add_table(rows=1, cols=2)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False
                table.columns[0].width = Inches(4.5)
                table.columns[1].width = Inches(2.37)
                
                # 左侧：公司/项目名 + 职位/技术栈
                cell_l = table.cell(0, 0)
                p_l = cell_l.paragraphs[0]
                p_l.paragraph_format.space_before = Pt(3)
                p_l.paragraph_format.space_after = Pt(1)
                r_name = p_l.add_run(item.get("name", ""))
                r_name.font.bold = True
                r_name.font.size = Pt(10)
                r_name.font.color.rgb = COLOR_TEXT
                r_name.font.name = "Microsoft YaHei"
                
                role_text = item.get("role", "") or item.get("tech_stack", "")
                if role_text:
                    r_role = p_l.add_run(f"  |  {role_text}")
                    r_role.font.size = Pt(9.5)
                    r_role.font.color.rgb = COLOR_PRIMARY
                    r_role.font.name = "Microsoft YaHei"
                
                # 右侧：时间
                cell_r = table.cell(0, 1)
                p_r = cell_r.paragraphs[0]
                p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_r.paragraph_format.space_before = Pt(3)
                p_r.paragraph_format.space_after = Pt(1)
                r_date = p_r.add_run(item.get("date", ""))
                r_date.font.size = Pt(9.5)
                r_date.font.color.rgb = COLOR_MUTED
                r_date.font.name = "Microsoft YaHei"
                
                # Bullet points
                for bullet in item.get("bullets", []):
                    p_b = doc.add_paragraph()
                    p_b.paragraph_format.left_indent = Inches(0.18)
                    p_b.paragraph_format.space_before = Pt(1)
                    p_b.paragraph_format.space_after = Pt(2)
                    p_b.paragraph_format.line_spacing = 1.15
                    
                    r_dot = p_b.add_run("•  ")
                    r_dot.font.size = Pt(9)
                    r_dot.font.color.rgb = COLOR_ACCENT
                    
                    r_txt = p_b.add_run(bullet)
                    r_txt.font.size = Pt(9.5)
                    r_txt.font.color.rgb = COLOR_TEXT
                    r_txt.font.name = "Microsoft YaHei"

    doc.save(output_path)
    print(f"Resume generated successfully at: {output_path}")

if __name__ == "__main__":
    # 测试样例数据
    sample_data = {
        "name": "张小凡",
        "target_role": "后端架构师 / 资深 Golang 工程师",
        "phone": "138-0000-0000",
        "email": "zhangxiaofan@example.com",
        "location": "北京 / 深圳",
        "links": "github.com/murenyu1",
        "education": [
            {
                "school": "北京航空航天大学",
                "major": "计算机科学与技术",
                "degree": "工学硕士",
                "date": "2019.09 - 2022.06"
            }
        ],
        "skills": [
            {"category": "后端语言与框架", "detail": "Golang, Python, C++, Gin, gRPC, Protobuf"},
            {"category": "存储与缓存", "detail": "MySQL (分库分表与慢查询调优), Redis (分布式锁/缓存架构), TiDB"},
            {"category": "中间件与系统", "detail": "Kafka, RabbitMQ, Docker, Kubernetes, Linux 内核性能调优"}
        ],
        "experience": [
            {
                "name": "知名互联网独角兽",
                "role": "资深后端开发工程师",
                "date": "2022.07 - 至今",
                "bullets": [
                    "主导交易核心微服务架构升级，重塑订单处理状态机，在高并发秒杀场景下将接口响应时间降低 70%，支撑峰值 QPS 5000+。",
                    "针对分布式事务一致性痛点，设计基于消息队列与最大努力通知方案，彻底杜绝数据长周期不一致事故，提升履约效率 15%。"
                ]
            }
        ],
        "projects": [
            {
                "name": "高可用分布式任务调度系统",
                "tech_stack": "Golang, Etcd, Raft, gRPC",
                "date": "2023.03 - 2023.10",
                "bullets": [
                    "基于 Raft 协议实现调度节点高可用选举与状态机复制，单集群支持日调度任务量 1000 万次以上。",
                    "针对节点宕机引起的任务丢失问题，研发自动心跳探测与故障自动漂移接管机制，故障恢复时间控制在 3 秒以内。"
                ]
            }
        ]
    }
    output_file = sys.argv[1] if len(sys.argv) > 1 else "sample_resume.docx"
    create_resume_docx(sample_data, output_file)
