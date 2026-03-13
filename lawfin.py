# 파일명: lawfin.py
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from streamlit_agraph import agraph, Node, Edge, Config
import io
import datetime
import time

# ==========================================
# [Class 1] Legal Ontology (사고 대응용 - V3 Refined for Sejong Persona)
# ==========================================
class LegalOntology:
    def __init__(self):
        # 세종 페르소나: 단순 법령 나열이 아닌, 실질적 해결책과 경영진 보호 논리 제공
        self.regulations = [
            {
                "condition": lambda c: c['type'] in ["횡령", "배임"] and c['amount'] >= 300_000_000,
                "id": "REG_001", "risk_level": "CRITICAL",
                "title": "특정경제범죄 가중처벌법 및 감독규정 위반",
                "law": "특정경제범죄 가중처벌 등에 관한 법률, 금융기관 검사 및 제재에 관한 규정 시행세칙 제67조(금융사고 보고)",
                "report_deadline": "즉시 (지체 없이)",
                "internal_rule": "내부통제규정 제12조 (사고 보고 및 직무 배제)",
                "precedents": ["A은행 700억 횡령 건: 내부통제 부실로 인한 경영진 중징계 사례", "B저축은행: 즉시 보고 위반 과태료 부과 사례"],
                "action_plan": {
                    "인사팀": "행위자 및 결재 라인 즉시 대기발령 및 직무 정지 (증거 인멸 방지)",
                    "법무팀": "형사 고소장 접수 및 피의자 자산 가압류 신청 (채권 보전)",
                    "감사팀": "자금 흐름 전수 조사 및 내부 공모자 파악 (Forensic 감사)",
                    "홍보팀": "언론 대응 시나리오 가동 (Reputation Risk 관리)"
                },
                "responsibility_map": {
                    "담당 임원": "CCO (준법감시인) 및 소관 본부장",
                    "CEO 리스크": "내부통제 기준 마련 의무 위반 여부에 따른 해임 권고 등 중징계 가능성 존재 (지배구조법 개정안 반영 필요)"
                }
            },
            {
                "condition": lambda c: c['dept'] == "IT" or "해킹" in c['desc'] or "권한" in c['desc'] or c['type'] == "정보유출",
                "id": "REG_002", "risk_level": "HIGH",
                "title": "전자금융거래법 및 개인정보보호법 위반",
                "law": "전자금융거래법 제21조(안전성의 확보의무), 신용정보법 제39조의4",
                "report_deadline": "인지 후 24시간 이내 (금융보안원/금감원)",
                "internal_rule": "정보보호지침 제2장 (접근통제 및 권한 관리)",
                "precedents": ["C증권사 홈트레이딩 지연 과태료 5천만원", "D카드사 정보유출 CEO 경고"],
                "action_plan": {
                    "CISO": "금융보안원 침해사고 신고 및 원인 분석 (로그 보존 필수)",
                    "IT운영팀": "취약점 패치, 비인가 접속 IP 차단, 계정 동결",
                    "법무팀": "이용자 통지문 작성 및 피해 구제 절차 검토"
                },
                "responsibility_map": {
                    "담당 임원": "CISO (정보보호최고책임자)",
                    "CEO 리스크": "정보보호 예산/인력 지원 적정성 입증 실패 시 관리 책임 소급 (전자금융감독규정)"
                }
            }
        ]

# ==========================================
# [Class 2] Execution Ontology (기술 검증용 - V4)
# ==========================================
class ExecutionOntology:
    def __init__(self):
        self.common_standards = {
            "클라우드 이용지원": {
                "checklist": ["VPC 분리", "Subnet 설계", "보안그룹(ACL)", "KMS 키관리", "로그 무결성"],
                "regulator_focus": "금융보안원 CSP 안전성 평가 기준 및 데이터 비식별화 조치",
                "risk_comment": "아키텍처 구성도만으로는 부족합니다. 실제 데이터 흐름 제어(ACL) 명세와 망분리 예외 승인 논리가 핵심입니다."
            },
            "혁신금융서비스 지정": {
                "checklist": ["처리속도(TPS)", "오류율 관리", "비상시망분리", "AI 설명가능성", "소비자오인방지"],
                "regulator_focus": "혁신성(정량지표), 소비자 보호 방안, 배상책임 이행 능력",
                "risk_comment": "법적 논리는 타당하나, 금융위 심사역들은 정량적 혁신성(Before/After)과 구체적인 소비자 보호 장치를 집중 질의할 것입니다."
            },
            "전자금융 기반시설": {
                "checklist": ["망분리 예외", "접근통제(2FA)", "DR센터(RTO/RPO)", "백업 소산", "침해사고대응"],
                "regulator_focus": "전자금융감독규정 제15조(해킹방지대책) 및 제3자 리스크 관리",
                "risk_comment": "단순 DR 센터 구축 여부보다, 실제 재해 시나리오별 전환 훈련(Mock Drill) 결과 보고서가 필수적입니다."
            }
        }

    def evaluate_readiness(self, sector, task_type, user_inputs):
        standard = self.common_standards.get(task_type)
        missing_items = []
        for item in standard['checklist']:
            if item not in user_inputs['technical_detail']:
                missing_items.append(item)
        score = 100 - (len(missing_items) * 20)
        if score < 0: score = 0
        status = "승인 가능" if score >= 80 else "보완 필요(반려 위험)"
        return {"score": score, "status": status, "missing": missing_items, 
                "consultant_view": standard['risk_comment'], "regulator_standard": standard['regulator_focus']}

# ==========================================
# [Helper] Visualization & Reporting Functions
# ==========================================
def draw_ontology_graph(case_data, findings):
    nodes = []
    edges = []
    
    # 금액 포맷팅
    formatted_amount = "{:,}".format(case_data['amount'])
    
    event_node_id = "Event_001"
    nodes.append(Node(id=event_node_id, label=f"🚨 사고: {case_data['type']}\n(규모: {formatted_amount}원)", size=35, color="#FF4B4B"))
    
    for idx, item in enumerate(findings):
        reg_id = f"Law_{idx}"
        nodes.append(Node(id=reg_id, label=item['title'], size=25, color="#1E88E5"))
        edges.append(Edge(source=event_node_id, target=reg_id, label="위반 법령"))
        
        resp_id = f"Resp_{idx}"
        nodes.append(Node(id=resp_id, label=f"👔 책임: {item['responsibility_map']['담당 임원']}", size=20, color="#9C27B0"))
        edges.append(Edge(source=reg_id, target=resp_id, label="귀속"))
        
        if "CEO" in item['responsibility_map']['CEO 리스크']:
            ceo_id = f"CEO_{idx}"
            nodes.append(Node(id=ceo_id, label="🏛️ CEO 제재 리스크", size=15, color="#FF9800"))
            edges.append(Edge(source=resp_id, target=ceo_id, label="확장"))
            
        for dept, action in item['action_plan'].items():
            act_id = f"Act_{idx}_{dept}"
            nodes.append(Node(id=act_id, label=f"✅ {dept}\n{action}", size=15, color="#4CAF50"))
            edges.append(Edge(source=reg_id, target=act_id, label="대응 Action"))
            
    config = Config(width="100%", height=600, directed=True, physics=True, hierarchical=False)
    return nodes, edges, config

def generate_sejong_report(case_data, analysis_results): 
    # 날짜 계산
    today = datetime.datetime.now()
    detection_time = today - datetime.timedelta(hours=2) # 2시간 전 인지 가정
    
    doc = Document()
    
    # 1. 헤더 (세종 스타일)
    title = doc.add_heading('[법무법인 세종] 금융사고 대응 및 법률 검토 보고서', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"수신: 귀사 경영진 및 컴플라이언스 위원회 귀중")
    doc.add_paragraph(f"일자: {today.strftime('%Y년 %m월 %d일')}")
    doc.add_paragraph("-" * 80)

    # 2. 사고 개요 (감독기관 보고 양식 준용)
    doc.add_heading('1. 사고 개요 (감독기관 보고용 요약)', level=1)
    
    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'
    
    # 테이블 내용 채우기
    rows = table.rows
    rows[0].cells[0].text = "사고 유형"
    rows[0].cells[1].text = case_data['type']
    
    rows[1].cells[0].text = "발생 일시"
    rows[1].cells[1].text = today.strftime('%Y-%m-%d %H:%M') + " (추정)"
    
    rows[2].cells[0].text = "인지 시점"
    rows[2].cells[1].text = detection_time.strftime('%Y-%m-%d %H:%M')
    
    rows[3].cells[0].text = "사고 금액"
    rows[3].cells[1].text = f"{case_data['amount']:,} 원"
    
    rows[4].cells[0].text = "행위자 / 소속"
    rows[4].cells[1].text = f"성명 불상 (내부 직원 추정) / {case_data['dept']}"
    
    rows[5].cells[0].text = "사고 내용"
    rows[5].cells[1].text = case_data['desc']
    
    # 법적 보고 시한 계산 (단순 예시)
    rows[6].cells[0].text = "보고 시한"
    reporting_limit = "즉시 (금융사고 보고규정)" if analysis_results and analysis_results[0]['risk_level'] == "CRITICAL" else "인지 후 24시간 이내"
    rows[6].cells[1].text = reporting_limit

    doc.add_paragraph("\n")

    # 3. 법률적 검토 의견 (세종 전문위원/변호사 페르소나)
    doc.add_heading('2. 법무법인 세종 법률 검토 의견', level=1)
    
    for item in analysis_results:
        p_title = doc.add_paragraph()
        runner = p_title.add_run(f"■ 관련 법령 및 규제: {item['title']}")
        runner.bold = True
        runner.font.color.rgb = RGBColor(0, 51, 102) # Navy Blue
        
        doc.add_paragraph(f"위반 소지 법령: {item['law']}", style='List Bullet')
        doc.add_paragraph(f"내부 규정 대조: {item['internal_rule']}", style='List Bullet')
        
        doc.add_heading("전문가 어드바이스 (Insight)", level=3)
        risk_text = f"본 건은 '{item['risk_level']}' 등급의 리스크로 판단됩니다. {item['precedents'][0]} 등 감독당국의 제재 기조를 볼 때, 단순 행위자 처벌을 넘어 경영진의 관리 감독 소홀(지배구조법상 내부통제 의무)로 확대될 가능성이 높습니다."
        doc.add_paragraph(risk_text)
        doc.add_paragraph(f"특히 CEO 리스크와 관련하여: {item['responsibility_map']['CEO 리스크']}")
        
        doc.add_heading("즉시 실행 제언 (Action Plan)", level=3)
        for dept, action in item['action_plan'].items():
            doc.add_paragraph(f"[{dept}] {action}", style='List Number')

    doc.add_paragraph("\n")
    
    # 4. 담당 변호사 및 긴급 연락처
    doc.add_heading('3. 담당 전문가 및 긴급 연락처', level=1)
    contact_table = doc.add_table(rows=2, cols=2)
    contact_table.style = 'Table Grid'
    
    contact_rows = contact_table.rows
    contact_rows[0].cells[0].text = "금융규제그룹장 (Partner)"
    contact_rows[0].cells[1].text = "변호사 정세종 (02-740-0001, sj.jung@sejong_fake.com)\n前 금융위원회 법률자문관"
    
    contact_rows[1].cells[0].text = "디지털포렌식 전문위원"
    contact_rows[1].cells[1].text = "전문위원 박보안 (010-1234-5678, hotlines@sejong_fake.com)\n前 금융감독원 IT검사국 팀장"

    doc.add_paragraph("\n※ 본 문서는 법률적 검토 의견서로, 감독기관 제출 전 반드시 담당 변호사의 최종 확인을 받으시기 바랍니다.")

    bio = io.BytesIO()
    doc.save(bio)
    return bio

def generate_gap_report(sector, task, result, inputs): # V4 Report
    doc = Document()
    doc.add_heading(f'[{sector}] {task} 기술 실행 적정성 검토', 0)
    doc.add_paragraph(f"검토자: 법무법인 세종 디지털금융센터")
    doc.add_heading('1. 진단 요약', level=1)
    doc.add_paragraph(f"현재 점수: {result['score']}점 ({result['status']})")
    doc.add_paragraph(f"전문가 의견: {result['consultant_view']}")
    doc.add_heading('2. 보완 필요 사항 (Execution Gap)', level=1)
    if result['missing']:
        for item in result['missing']:
            doc.add_paragraph(f"❌ [누락] {item} : 구체적 명세 필요", style='List Bullet')
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# ==========================================
# [Main UI]
# ==========================================
def main():
    st.set_page_config(page_title="Sejong Legal-Tech Platform", page_icon="⚖️", layout="wide")
    
    # 로고 및 헤더 (가상)
    st.title("⚖️ 법무법인 세종 | Financial Risk Guardian")
    st.markdown("**금융규제 · 디지털 혁신 · 리스크 관리 통합 솔루션**")
    st.caption("Designed by Digital Finance Group (Former Regulators & Top-tier Experts)")
    st.markdown("---")

    # 탭 분리 (3번째 탭 추가)
    tab1, tab2, tab3 = st.tabs(["🚨 [1] 금융사고 대응 & 리스크 시각화", "🏗️ [2] 디지털 신사업 기술/규제 검증", "💬 [3] 로펌 채팅 상담"])

    # ---------------------------------------------------------
    # TAB 1: 사고 대응
    # ---------------------------------------------------------
    with tab1:
        st.subheader("🕸️ 금융사고 긴급 대응 및 규제 리스크 분석")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.info("💡 사고 정황을 입력하시면, 세종의 금융 전문 변호사와 전문위원의 식견이 담긴 분석 리포트를 생성합니다.")
            
            accident_type = st.selectbox("사고 유형", ["선택하세요", "횡령", "배임", "정보유출", "시스템장애"])
            actor_dept = st.text_input("행위자/발생 부서", "IT운영팀")
            
            # 금액 입력
            accident_amount = st.number_input("피해 추정 금액 (원)", min_value=0, value=500000000, step=1000000, format="%d")
            st.caption(f"입력 금액: {accident_amount:,} 원") # 사용자가 보기 편하게 캡션 추가
            
            accident_desc = st.text_area("사고 내용 (경위 상세)", "관리자(root) 계정을 도용하여 고객 휴면 계좌 잔액 5억 원을 차명 계좌로 이체하고 접속 로그 삭제를 시도함.", height=150)
            
            analyze_btn_v3 = st.button("🚀 사고 분석 및 대응 리포트 생성")

        with col2:
            if analyze_btn_v3 and accident_type != "선택하세요":
                case_data = {"type": accident_type, "dept": actor_dept, "amount": accident_amount, "desc": accident_desc}
                engine = LegalOntology()
                findings = []
                for rule in engine.regulations:
                    if rule['condition'](case_data):
                        findings.append(rule)
                
                if findings:
                    # 1. 보고서 다운로드 섹션
                    st.success(f"✅ 분석 완료: {len(findings)}건의 중대 규제 리스크가 식별되었습니다.")
                    st.markdown("### 📄 사고 대응 보고서 (Regulator Submission Draft)")
                    st.markdown("금융감독원 등 감독기관 보고 양식을 준수하며, 경영진을 위한 법률적 방어 논리가 포함되어 있습니다.")
                    
                    doc_file_sejong = generate_sejong_report(case_data, findings)
                    st.download_button(
                        label="📥 [법무법인 세종] 금융사고 대응 보고서 다운로드 (Word)",
                        data=doc_file_sejong.getvalue(),
                        file_name=f"Sejong_Accident_Report_{datetime.datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    st.markdown("---")
                    
                    # 2. 시각화 그래프
                    st.markdown("### 🕸️ 리스크 & 책임 관계도 (Risk & Responsibility Graph)")
                    nodes, edges, config = draw_ontology_graph(case_data, findings)
                    agraph(nodes=nodes, edges=edges, config=config)
                    
                else:
                    st.info("입력하신 조건에서는 특이 중대 리스크가 식별되지 않았습니다. (일반 절차 준수 요망)")

    # ---------------------------------------------------------
    # TAB 2: 기술 실행 검증
    # ---------------------------------------------------------
    with tab2:
        st.subheader("🛠️ 디지털 신사업 기술/규제 실행 검증 (Execution Readiness)")
        st.markdown("> **EU 및 글로벌 규제 동향을 반영하여, 국내 금융기관이 과도한 부담 없이 규제를 준수하며 혁신할 수 있는 현실적 대안을 제시합니다.**")
        
        col_exec_1, col_exec_2 = st.columns([1, 1])
        
        with col_exec_1:
            st.markdown("### 1. 고객사 및 업무 정의")
            target_sector = st.radio("대상 산업군", ("은행 (Bank)", "증권 (Securities)", "카드 (Card)", "보험 (Insurance)"), horizontal=True)
            task_type = st.selectbox("검토 대상 업무", ["클라우드 이용지원", "혁신금융서비스 지정", "전자금융 기반시설"])
            project_name = st.text_input("프로젝트명", f"{target_sector.split()[0]} 차세대 시스템 구축")
            
            st.markdown("### 2. 현업 기술 명세 입력")
            tech_input = st.text_area("기술 아키텍처/보안 명세", height=150, placeholder="AWS 도입 예정, 망분리 준수함. (구체적인 기술 키워드: VPC, KMS, ACL 등을 입력해보세요)")
            check_btn_v4 = st.button("🔍 기술 실행 적합성(Gap) 진단")

        with col_exec_2:
            if check_btn_v4:
                engine_v4 = ExecutionOntology()
                inputs = {"project_name": project_name, "technical_detail": tech_input}
                
                with st.spinner(f"[{target_sector}] 규제 기준 대조 중..."):
                    time.sleep(1)
                    result = engine_v4.evaluate_readiness(target_sector, task_type, inputs)
                
                st.markdown("### 3. 진단 결과")
                score_c1, score_c2 = st.columns([1, 2])
                with score_c1:
                    st.metric("규제 이행 완성도", f"{result['score']}점")
                with score_c2:
                    if result['score'] >= 80: st.success(result['status'])
                    else: st.error(result['status'])
                
                st.info(f"💬 **세종 전문위원 코멘트:** {result['consultant_view']}")
                
                if result['missing']:
                    st.warning("**[반려 위험] 누락된 핵심 기술 요소 (Execution Gap):**")
                    for item in result['missing']:
                        st.write(f"- ❌ {item}")
                    st.caption("👉 위 항목은 감독기관 심사 시 필수 체크리스트입니다. 보완이 필요합니다.")
                else:
                    st.success("✅ 기술적 실행 요건이 규제 수준을 충족합니다.")

                st.markdown("---")
                doc_file_v4 = generate_gap_report(target_sector, task_type, result, inputs)
                st.download_button("📥 [기술] 실행 보완 보고서 다운로드", doc_file_v4.getvalue(), f"Execution_Report_{target_sector}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # ---------------------------------------------------------
    # TAB 3: 채팅 상담 (신규 추가)
    # ---------------------------------------------------------
    with tab3:
        st.subheader("💬 1:1 법률 전문가 채팅 상담")
        
        # 안내 문구 (요청 사항)
        st.info("📢 **채팅을 남겨 놓으시면 담당 변호사가 바로 연락 드리도록 하겠습니다.**")

        # 채팅 히스토리 초기화
        if "messages" not in st.session_state:
            st.session_state.messages = []

        # 이전 대화 내용 표시
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # 사용자 입력 처리
        if prompt := st.chat_input("상담하실 내용을 입력해주세요... (예: 마이데이터 사업 인가 요건, 망분리 예외 승인 절차 등)"):
            # 1. 사용자 메시지 표시
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            # 2. 어시스턴트(세종) 응답 시뮬레이션
            time.sleep(1) # 생각하는 척 지연 효과
            response_text = f"문의해 주셔서 감사합니다. 귀하께서 남겨주신 내용('{prompt}')은 접수되었습니다.\n\n본 사안에 가장 적합한 **금융규제그룹 담당 변호사**가 배정되어 24시간 이내에 연락드리겠습니다."
            
            st.session_state.messages.append({"role": "assistant", "content": response_text})
            with st.chat_message("assistant"):
                st.markdown(response_text)

if __name__ == "__main__":
    main()
