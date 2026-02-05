import streamlit as st
import pandas as pd
from docx import Document
import os
from datetime import datetime
import io
import re

# -------------------- 页面配置 --------------------
st.set_page_config(
    page_title="保密协议生成器 | 国联新创",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="expanded"
)

# -------------------- 自定义CSS样式 --------------------
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 0.5rem;
        padding-top: 1rem;
    }
    .sub-header {
        text-align: center;
        color: #64748B;
        margin-bottom: 2rem;
        font-size: 1.1rem;
    }
    .stButton>button {
        background-color: #3B82F6;
        color: white;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #2563EB;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2);
    }
    .success-box {
        padding: 1.5rem;
        border-radius: 0.5rem;
        background-color: #D1FAE5;
        border: 1px solid #10B981;
        margin: 1.5rem 0;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #EFF6FF;
        border: 1px solid #3B82F6;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #FEF3C7;
        border: 1px solid #F59E0B;
        margin: 1rem 0;
    }
    .step-box {
        background-color: #F8FAFC;
        border-left: 4px solid #3B82F6;
        padding: 1rem;
        margin-bottom: 1.5rem;
        border-radius: 0 0.5rem 0.5rem 0;
    }
    .company-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 0.5rem;
        padding: 1.25rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    }
</style>
""", unsafe_allow_html=True)


# -------------------- 核心函数定义 --------------------

def load_template():
    """加载并验证Word模板文件"""
    template_path = "保密协议模板.docx"

    if not os.path.exists(template_path):
        st.error(f"❌ **关键错误**：未找到模板文件 '{template_path}'")
        st.info("""
        **解决方法：**
        1. 请将您的《保密协议模板.docx》文件放在与此程序相同的目录下
        2. 确保模板中包含以下精确的占位符文本：
           - `[千寻智能(杭州)科技有限公司]`
           - `[浙江省杭州市萧山区宁围街道利一路188号天人大厦浙大研究院数字经济孵化器4层401室-38]`
        """)
        st.stop()

    # 验证模板中是否包含必要的占位符
    try:
        doc = Document(template_path)
        full_text = "\n".join([para.text for para in doc.paragraphs])

        required_placeholders = [
            "[千寻智能(杭州)科技有限公司]",
            "[浙江省杭州市萧山区宁围街道利一路188号天人大厦浙大研究院数字经济孵化器4层401室-38]"
        ]

        missing = []
        for placeholder in required_placeholders:
            if placeholder not in full_text:
                missing.append(placeholder)

        if missing:
            st.error(f"❌ **模板验证失败**：模板中缺少以下占位符：")
            for m in missing:
                st.code(m, language="text")
            st.info("请在模板文件中添加上述占位符，然后重启应用。")
            st.stop()

        return template_path
    except Exception as e:
        st.error(f"读取模板文件时出错：{str(e)}")
        st.stop()


def smart_replace_in_document(doc, replace_pairs):
    """
    智能替换文档中的文本（增强版）
    处理跨多个Run的文本替换问题
    """
    # 1. 替换所有段落
    for para in doc.paragraphs:
        original_text = para.text
        new_text = original_text

        # 对每个占位符进行替换
        for old, new in replace_pairs.items():
            if old in new_text:
                new_text = new_text.replace(old, new)

        # 如果文本发生了变化，更新段落
        if new_text != original_text:
            # 清空所有runs
            for run in para.runs:
                run.text = ""
            # 重新设置文本到第一个run
            if para.runs:
                para.runs[0].text = new_text
            else:
                # 如果没有run，添加一个
                para.add_run(new_text)

    # 2. 替换表格中的文本
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    original_text = para.text
                    new_text = original_text

                    for old, new in replace_pairs.items():
                        if old in new_text:
                            new_text = new_text.replace(old, new)

                    if new_text != original_text:
                        for run in para.runs:
                            run.text = ""
                        if para.runs:
                            para.runs[0].text = new_text

    return doc


def mock_query_company_address(company_name):
    """
    模拟查询公司地址
    注意：这是一个演示函数。实际使用时需要接入企查查/天眼查等商业API
    """
    # 模拟数据 - 可以在这里添加您常用的公司
    mock_database = {
        "千寻智能(杭州)科技有限公司": "浙江省杭州市萧山区宁围街道利一路188号天人大厦浙大研究院数字经济孵化器4层401室-38",
        "苏州易航智能科技有限公司": "江苏省苏州市苏州工业园区金鸡湖大道88号人工智能产业园G1栋",
        "深圳元宇互动科技有限公司": "广东省深圳市南山区粤海街道科苑路8号科技大厦西座12楼1201室",
        "北京智云科技有限公司": "北京市海淀区中关村大街1号鼎好大厦A座12层",
        "上海未来机器人有限公司": "上海市浦东新区张江高科技园区科苑路151号"
    }

    # 尝试模糊匹配（如果公司名不完全一致）
    for key, address in mock_database.items():
        if company_name in key or key in company_name:
            return address

    # 完全匹配
    return mock_database.get(company_name, None)


def safe_filename(text, max_length=50):
    """
    生成安全的文件名，移除不安全的字符
    修复了正则表达式错误
    """
    # 修复后的正则表达式：允许字母、数字、下划线、空格、括号、连字符和中文
    # 注意：将连字符-放在字符类的最后，避免被解释为范围
    safe_text = re.sub(r'[^\w\s()（）\-]', '', text)

    # 移除多余的空格
    safe_text = re.sub(r'\s+', ' ', safe_text).strip()

    # 限制长度
    return safe_text[:max_length]


def generate_document(company_name, company_address, template_path):
    """生成新的保密协议文档"""
    try:
        # 加载模板
        doc = Document(template_path)

        # 定义替换规则
        replace_pairs = {
            "[千寻智能(杭州)科技有限公司]": company_name,
            "[浙江省杭州市萧山区宁围街道利一路188号天人大厦浙大研究院数字经济孵化器4层401室-38]": company_address
        }

        # 执行智能替换
        doc = smart_replace_in_document(doc, replace_pairs)

        # 将文档保存到内存字节流
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)  # 将指针移回文件开头

        return file_stream, None  # 返回文档流和错误信息（无错误）

    except Exception as e:
        return None, str(e)


# -------------------- 主应用界面 --------------------

def main():
    # 标题区域
    st.markdown('<h1 class="main-header">📄 保密协议智能生成器</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">无锡国联新创私募投资基金有限公司 · 内部工具</p>', unsafe_allow_html=True)

    # 加载并验证模板
    template_path = load_template()

    # 创建两列布局
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("### 填写协议信息")

        # 步骤1：公司信息
        with st.container():
            st.markdown('<div class="step-box"><strong>步骤1：输入公司信息</strong></div>', unsafe_allow_html=True)

            company_name = st.text_input(
                "**目标公司全称** *",
                placeholder="请输入与营业执照一致的完整公司名称",
                help="请务必确保公司名称准确无误，它将直接填入协议中。",
                key="company_name"
            )

            if not company_name:
                st.info("👆 请输入公司名称以继续")
                st.stop()

        # 步骤2：地址获取
        with st.container():
            st.markdown('<div class="step-box"><strong>步骤2：获取公司地址</strong></div>', unsafe_allow_html=True)

            # 地址获取方式选择
            address_mode = st.radio(
                "**选择地址获取方式：**",
                ["🔍 尝试自动查询", "✏️ 手动填写地址"],
                index=1,  # 默认选手动填写
                horizontal=True,
                key="address_mode"
            )

            company_address = ""

            if address_mode == "🔍 尝试自动查询":
                st.markdown(
                    '<div class="info-box">💡 当前为模拟查询模式，仅支持有限的演示数据。如需真实查询，需接入企业信息API。</div>',
                    unsafe_allow_html=True)

                if st.button("🚀 点击查询公司地址", use_container_width=True, type="secondary"):
                    with st.spinner("正在查询公司地址..."):
                        # 模拟查询
                        company_address = mock_query_company_address(company_name)

                        if company_address:
                            st.success(f"✅ 查询成功！")
                            st.markdown(
                                f'<div class="company-card"><strong>公司名称：</strong>{company_name}<br><strong>注册地址：</strong>{company_address}</div>',
                                unsafe_allow_html=True)
                        else:
                            st.warning("未找到该公司地址。")
                            st.info("""
                            **可能原因：**
                            1. 公司名称与模拟数据库不匹配
                            2. 当前为演示模式，数据有限

                            **建议：** 切换到"手动填写地址"方式
                            """)

                # 如果查询失败或未查询，显示手动输入框作为备用
                if not company_address:
                    st.divider()
                    st.markdown("**或直接手动填写地址：**")
                    company_address = st.text_area(
                        "公司注册地址",
                        placeholder="请准确填写公司的工商注册地址，格式：省 市 区 街道 门牌号 楼层/房间号",
                        height=120,
                        key="manual_address_backup",
                        help="此地址将直接填入协议中，请仔细核对。"
                    )
            else:
                # 手动填写模式
                company_address = st.text_area(
                    "**公司注册地址** *",
                    placeholder="请准确填写公司的工商注册地址，格式：省 市 区 街道 门牌号 楼层/房间号",
                    height=120,
                    key="manual_address",
                    help="此地址将直接填入协议中，请仔细核对。"
                )

        # 只有获取到地址后才显示生成按钮
        if company_address:
            st.markdown('<div class="step-box"><strong>步骤3：生成协议文档</strong></div>', unsafe_allow_html=True)

            # 信息预览
            with st.expander("📋 预览生成信息", expanded=True):
                preview_col1, preview_col2 = st.columns(2)
                with preview_col1:
                    st.metric("公司名称", company_name[:20] + "..." if len(company_name) > 20 else company_name)
                with preview_col2:
                    st.metric("地址长度", f"{len(company_address)} 字符")

                st.caption("完整地址预览：")
                st.info(company_address)

            # 生成按钮
            if st.button("🎯 生成保密协议文件", type="primary", use_container_width=True):
                with st.spinner("正在生成协议文档，请稍候..."):
                    # 调用生成函数
                    file_stream, error = generate_document(company_name, company_address, template_path)

                    if error:
                        st.error(f"生成文档时出错：{error}")
                        st.info("""
                        **常见问题排查：**
                        1. 请检查模板文件是否被其他程序打开
                        2. 确保模板文件格式正确（.docx格式）
                        3. 重启应用后重试
                        """)
                    else:
                        # 使用安全的文件名生成函数
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        safe_name = safe_filename(company_name, 50)
                        download_name = f"保密协议_{safe_name}_{timestamp}.docx"

                        # 显示成功消息
                        st.markdown(f"""
                        <div class="success-box">
                            <h4>✅ 文档生成成功！</h4>
                            <p><strong>文件名称：</strong> {download_name}</p>
                            <p><strong>生成时间：</strong> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                            <p>请点击下方按钮下载文档。下载后请仔细检查文档内容，特别是公司名称和地址的准确性。</p>
                        </div>
                        """, unsafe_allow_html=True)

                        # 提供下载按钮
                        st.download_button(
                            label="📥 下载保密协议文档",
                            data=file_stream,
                            file_name=download_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="primary"
                        )

                        # 操作建议
                        st.divider()
                        st.caption("💡 **下一步建议**：下载并核对文档后，可以：")
                        st.caption("1. 直接打印使用")
                        st.caption("2. 如需生成另一份协议，请刷新页面或修改上方信息")

                        # 成功日志（可选）
                        st.session_state.last_generated = {
                            "company": company_name,
                            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "filename": download_name
                        }

    with col2:
        st.markdown("### 📖 使用指南")

        with st.expander("操作步骤", expanded=True):
            st.markdown("""
            1. **输入公司名称**  
               完整、准确的公司全称

            2. **获取公司地址**  
               - 自动查询：演示模式，数据有限  
               - 手动填写：最可靠的方式

            3. **预览并生成**  
               核对信息后生成文档

            4. **下载使用**  
               保存到本地并仔细核对
            """)

        st.divider()

        st.markdown("### ⚙️ 模板状态")
        try:
            doc = Document(template_path)
            file_size = os.path.getsize(template_path) / 1024
            mod_time = datetime.fromtimestamp(os.path.getmtime(template_path))

            st.success(f"✅ 模板正常")
            st.caption(f"大小: {file_size:.1f} KB")
            st.caption(f"修改: {mod_time.strftime('%Y-%m-%d %H:%M')}")

            # 检查占位符
            full_text = "\n".join([para.text for para in doc.paragraphs[:5]])
            placeholders = [
                "[千寻智能(杭州)科技有限公司]",
                "[浙江省杭州市萧山区宁围街道利一路188号天人大厦浙大研究院数字经济孵化器4层401室-38]"
            ]

            found_all = all(p in full_text for p in placeholders)
            if found_all:
                st.success("✅ 所有占位符就绪")
            else:
                st.warning("⚠️ 请检查占位符")

        except Exception as e:
            st.error(f"❌ 模板异常: {str(e)}")

        st.divider()

        st.markdown("### 🗃️ 模拟数据公司")
        st.caption("自动查询可用的演示数据：")
        demo_companies = [
            "千寻智能(杭州)科技有限公司",
            "苏州易航智能科技有限公司",
            "深圳元宇互动科技有限公司",
            "北京智云科技有限公司",
            "上海未来机器人有限公司"
        ]

        for company in demo_companies:
            if st.button(f"📌 {company[:12]}...", key=f"demo_{company}", use_container_width=True):
                st.session_state.company_name = company
                st.rerun()

        st.divider()

        # st.markdown("### 🔧 技术支持")
        # st.caption("**遇到问题？**")
        # st.caption("1. 检查模板文件是否存在")
        # st.caption("2. 确保占位符格式正确")
        # st.caption("3. 重启应用尝试")
        #
        # if st.button("🔄 重启应用", use_container_width=True, type="secondary"):
        #     st.rerun()


# -------------------- 应用启动 --------------------
if __name__ == "__main__":
    # 初始化session state
    if "company_name" not in st.session_state:
        st.session_state.company_name = ""

    main()