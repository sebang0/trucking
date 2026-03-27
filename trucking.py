import streamlit as st
import pandas as pd
import itertools
import math
import time
import plotly.express as px
from io import BytesIO

# 메모리 에러 방지용 설정
pd.set_option("styler.render.max_elements", 5000000)

st.set_page_config(page_title="배차 효율화 및 비용 절감 대시보드", layout="wide", initial_sidebar_state="expanded")

# 💡 [요청 반영 1] 상단 타이틀 및 우측 PDF 인쇄 버튼 (HTML/JS 활용)
col_title, col_print = st.columns([8.5, 1.5])
with col_title:
    st.title("🚛 왕복 배차 / 자차 투입 전환 시뮬레이션")
with col_print:
    st.components.v1.html(
        """
        <button onclick="window.parent.print()" style="width: 100%; padding: 12px; margin-top: 15px; font-size: 16px; font-weight: bold; cursor: pointer; background-color: #f0f2f6; color: #31333F; border: 1px solid #c4c4c4; border-radius: 8px;">
        🖨️ PDF 인쇄
        </button>
        """,
        height=70
    )

# 💡 [요청 반영 2] 세션 상태 초기화 (버튼 클릭 여부 및 애니메이션 제어용)
if 'sim_run' not in st.session_state:
    st.session_state.sim_run = False
if 'show_anim' not in st.session_state:
    st.session_state.show_anim = False

# 기존 시/도 명칭 매핑 유지
KOR_COORDS = {
    '서울특별시': (37.5665, 126.9780), '서울': (37.5665, 126.9780),
    '부산광역시': (35.1796, 129.0756), '부산': (35.1796, 129.0756),
    '대구광역시': (35.8714, 128.6014), '대구': (35.8714, 128.6014),
    '인천광역시': (37.4563, 126.7052), '인천': (37.4563, 126.7052),
    '광주광역시': (35.1595, 126.8526), '광주': (35.1595, 126.8526),
    '대전광역시': (36.3504, 127.3845), '대전': (36.3504, 127.3845),
    '울산광역시': (35.5384, 129.3114), '울산': (35.5384, 129.3114),
    '세종특별자치시': (36.4800, 127.2890), '세종': (36.4800, 127.2890),
    '경기도': (37.2752, 127.0095), '경기': (37.2752, 127.0095),
    '강원도': (37.8854, 127.7298), '강원': (37.8854, 127.7298), '강원특별자치도': (37.8854, 127.7298),
    '충청북도': (36.6358, 127.4913), '충북': (36.6358, 127.4913),
    '충청남도': (36.6588, 126.6728), '충남': (36.6588, 126.6728),
    '전라북도': (35.8203, 127.1088), '전북': (35.8203, 127.1088), '전북특별자치도': (35.8203, 127.1088),
    '전라남도': (34.8161, 126.4629), '전남': (34.8161, 126.4629),
    '경상북도': (36.5760, 128.5056), '경북': (36.5760, 128.5056),
    '경상남도': (35.2383, 128.6922), '경남': (35.2383, 128.6922),
    '제주특별자치도': (33.4890, 126.4983), '제주': (33.4890, 126.4983)
}

# --- 핵심 로직 ---
@st.cache_data
def analyze_monthly_roi(df_orders, cost_dict, trips_per_month):
    df_orders['운송일자'] = pd.to_datetime(df_orders['운송일자'])
    df_orders = df_orders.sort_values('운송일자')
    orders_list = df_orders.to_dict('records')
    
    matched_ids = set() 
    round_trips = []
    
    outbound_lookup = {}
    for order in orders_list:
        key = (order['운송일자'], order['출발지'], order['도착지'], order.get('차종', '미상'))
        if key not in outbound_lookup: outbound_lookup[key] = []
        outbound_lookup[key].append(order)

    for in_order in orders_list:
        if in_order['오더번호'] in matched_ids: continue
        target_date = in_order['운송일자'] - pd.Timedelta(days=1)
        in_vehicle = in_order.get('차종', '미상')
        
        lookup_key = (target_date, in_order['도착지'], in_order['출발지'], in_vehicle)
        potential_matches = outbound_lookup.get(lookup_key, [])
        matching_outbound = [o for o in potential_matches if o['오더번호'] not in matched_ids]
        
        valid_combination, total_outbound_hire_cost, match_type = None, 0, ''
        
        if in_order['화종'] == '유해화학':
            chem_matches = [o for o in matching_outbound if o['화종'] == '유해화학' and o['중량'] <= in_order['중량']]
            if chem_matches:
                valid_combination = [chem_matches[0]['오더번호']]
                total_outbound_hire_cost = chem_matches[0]['용차단가(원)']
                match_type = '유해화학 단독'
        else:
            general_outbound = [o for o in matching_outbound if o['화종'] != '유해화학']
            max_weight = 0
            for r in range(1, min(len(general_outbound) + 1, 4)): 
                for combo in itertools.combinations(general_outbound, r):
                    total_weight = sum(item['중량'] for item in combo)
                    if max_weight < total_weight <= in_order['중량']:
                        max_weight = total_weight
                        valid_combination = [item['오더번호'] for item in combo]
                        total_outbound_hire_cost = sum(item['용차단가(원)'] for item in combo)
                        match_type = '일반 합적'

        if valid_combination:
            matched_ids.add(in_order['오더번호'])
            matched_ids.update(valid_combination)
            
            # 직행/경유 분리
            inbound_order_count = len(valid_combination)
            dispatch_type = '단독(직행)' if inbound_order_count == 1 else '합적(경유)'
            
            round_trips.append({
                '운행시작일(상행)': target_date.strftime('%Y-%m-%d'),
                '운행종료일(하행)': in_order['운송일자'].strftime('%Y-%m-%d'),
                '구간': f"{in_order['도착지']} ↔ {in_order['출발지']}",
                '지역1': sorted([in_order['도착지'], in_order['출발지']])[0],
                '지역2': sorted([in_order['도착지'], in_order['출발지']])[1],
                '상행_배차형태': dispatch_type,
                '차종': in_vehicle, '요구톤급(t)': float(in_order['중량']),
                '하행오더': in_order['오더번호'], '상행오더(리스트)': ', '.join(valid_combination),
                '상행금액(원)': total_outbound_hire_cost,
                '하행금액(원)': in_order['용차단가(원)'],
                '용차비_합계(원)': total_outbound_hire_cost + in_order['용차단가(원)']
            })

    monthly_summary = []
    rec_ids = set() 
    grouped_trips = {}
    
    for rt in round_trips:
        key = (rt['구간'], rt['지역1'], rt['지역2'], rt['차종'], rt['요구톤급(t)'], rt['상행_배차형태'])
        if key not in grouped_trips: grouped_trips[key] = []
        grouped_trips[key].append(rt)

    for key, trips in grouped_trips.items():
        route, loc1, loc2, v_type, ton, dispatch_type = key
        trip_count = len(trips)
        total_hire_cost_saved = sum(t['용차비_합계(원)'] for t in trips)
        
        req_trucks = math.ceil(trip_count / trips_per_month)
        cost_key = (loc1, loc2, v_type, ton)
        monthly_truck_cost = cost_dict.get(cost_key, 0)
        
        if monthly_truck_cost == 0:
            status = '왕복 배차 (원가 미상)'
            total_own_fleet_cost = 0
            monthly_saving = 0
        else:
            total_own_fleet_cost = req_trucks * monthly_truck_cost
            monthly_saving = total_hire_cost_saved - total_own_fleet_cost
            status = '자차 투입 (효율 달성)' if monthly_saving > 0 else '왕복 배차 (자차 비효율)'
            
        if '자차 투입' in status:
            for t in trips:
                rec_ids.add(t['하행오더'])
                rec_ids.update([o.strip() for o in t['상행오더(리스트)'].split(',')])
                
        monthly_summary.append({
            '투입_판단': status, '운행_구간': route, '요구_차종': v_type, '요구_톤급(t)': ton,
            '상행_배차형태': dispatch_type,
            '월_왕복횟수': trip_count, '필요_자차대수': req_trucks, '대당_월고정비(원)': monthly_truck_cost,
            '월_용차전환비(A)': total_hire_cost_saved, '월_자차운영비(B)': total_own_fleet_cost, '월_비용절감액(A-B)': monthly_saving
        })

    summary_df = pd.DataFrame(monthly_summary)
    if not summary_df.empty:
        summary_df['SortOrder'] = summary_df['투입_판단'].map({'자차 투입 (효율 달성)': 1, '왕복 배차 (자차 비효율)': 2, '왕복 배차 (원가 미상)': 3})
        summary_df = summary_df.sort_values(by=['SortOrder', '월_비용절감액(A-B)'], ascending=[True, False]).drop(columns=['SortOrder'])
        
    rt_df = pd.DataFrame(round_trips)
    if not rt_df.empty:
        rt_df = rt_df.drop(columns=['지역1', '지역2']) 
        
    return summary_df, rec_ids, matched_ids, rt_df

# --- 사이드바 설정 ---
st.sidebar.header("📁 데이터 및 조건 설정")
order_file = st.sidebar.file_uploader("1. 월간 운송 오더 엑셀", type=['xlsx'])
cost_file = st.sidebar.file_uploader("2. 마스터 자차 운행비 엑셀", type=['xlsx'])

st.sidebar.markdown("---")
target_trips_per_month = st.sidebar.slider("자차 1대당 월 목표 왕복 횟수", min_value=10, max_value=30, value=20)
rt_discount_rate = st.sidebar.number_input("왕복 배차 시 할인율 (%)", min_value=0, max_value=50, value=10, step=1)

st.sidebar.markdown("---")
# 💡 [요청 반영 2] 수동 실행 버튼
if st.sidebar.button("🚀 시뮬레이션 실행", type="primary", use_container_width=True):
    if order_file is not None and cost_file is not None:
        st.session_state.sim_run = True
        st.session_state.show_anim = True
    else:
        st.sidebar.error("엑셀 파일 2개를 모두 업로드해 주세요.")

# --- 메인 실행 블록 ---
if st.session_state.sim_run and order_file is not None and cost_file is not None:
    
    # 💡 [요청 반영 2] 실행 버튼을 눌렀을 때 1회만 트럭 진행률 애니메이션 표출
    if st.session_state.show_anim:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i in range(1, 101):
            # 트럭 이동 애니메이션 (스페이스바로 여백 추가)
            spaces = "&nbsp;" * (i // 2)
            status_text.markdown(f"**데이터 분석 및 최적화 매칭 진행 중... {i}%**<br><span style='font-size: 24px;'>{spaces}🚍</span>", unsafe_allow_html=True)
            progress_bar.progress(i)
            time.sleep(0.015) # 부드러운 애니메이션을 위한 딜레이
            
        status_text.empty()
        progress_bar.empty()
        st.session_state.show_anim = False # 다음 번 화면 새로고침(행 클릭 등) 시엔 뜨지 않음

    # 파일 포인터 초기화 및 데이터 읽기
    order_file.seek(0)
    cost_file.seek(0)
    raw_df = pd.read_excel(order_file)
    cost_df = pd.read_excel(cost_file)
    
    # 불량 데이터 필터링
    raw_df['출발지'] = raw_df['출발지'].astype(str).str.strip()
    raw_df['도착지'] = raw_df['도착지'].astype(str).str.strip()
    
    invalid_mask = raw_df['출발지'].isin(['0', '0.0', 'nan', '']) | raw_df['도착지'].isin(['0', '0.0', 'nan', ''])
    invalid_df = raw_df[invalid_mask]
    raw_df = raw_df[~invalid_mask]
    
    if not invalid_df.empty:
        st.warning(f"⚠️ 출발지 또는 도착지가 '0'이나 빈칸으로 입력된 불량 데이터 {len(invalid_df)}건이 집계에서 제외되었습니다.")
        with st.expander("❌ 제외된 불량 데이터 내역 표 확인 (클릭)"):
            st.dataframe(invalid_df, use_container_width=True, hide_index=True)
        st.divider()
    
    if raw_df.empty:
        st.error("정상 오더 데이터가 없습니다.")
        st.stop()
        
    cost_dict = {}
    for _, row in cost_df.iterrows():
        l1, l2 = sorted([str(row['지역1']).strip(), str(row['지역2']).strip()])
        key = (l1, l2, str(row['차종']), float(row['톤급(t)']))
        cost_dict[key] = row['자차운행비(원)']

    summary_df, deploy_ids, roundtrip_ids, rt_df = analyze_monthly_roi(raw_df, cost_dict, target_trips_per_month)
    
    # --- 1. 깔때기(Funnel) 요약 ---
    st.markdown("## 🎯 월간 배차 효율화 요약 (Funnel)")
    
    total_orders = len(raw_df)
    total_cost = raw_df['용차단가(원)'].sum()
    rt_count = len(roundtrip_ids)
    deploy_count = len(deploy_ids)
    
    deploy_truck_count = summary_df.loc[summary_df['투입_판단'] == '자차 투입 (효율 달성)', '필요_자차대수'].sum() if not summary_df.empty else 0
    
    col1, col2, col3 = st.columns(3)
    col1.metric("1️⃣ 전체 대상 물량", f"{total_orders:,} 건", f"총 용차비: {total_cost/1000000:,.0f} 백만원", delta_color="off")
    col2.metric("2️⃣ 왕복 배차 가능 (물리적)", f"{rt_count:,} 건", f"전환율: {rt_count/total_orders*100:.1f} %", delta_color="off")
    col3.metric("3️⃣ 자차 투입 확정 (효율 달성)", f"{deploy_count:,} 건 ({deploy_truck_count:,.0f} 대)", f"최종 채택률: {deploy_count/total_orders*100:.1f} %")

    st.divider()

    # --- 2. 원본 데이터 현황 ---
    st.markdown("### 1️⃣ 원본 데이터 현황 (전체 대상)")
    
    raw_df['시도_매칭키'] = raw_df['출발지'].apply(lambda x: str(x).split()[0][:2] if len(str(x).split()) > 0 else "")
    map_data = raw_df.groupby('시도_매칭키', as_index=False)['용차단가(원)'].sum()
    map_data['lat'] = map_data['시도_매칭키'].map(lambda x: KOR_COORDS.get(x, (None, None))[0])
    map_data['lon'] = map_data['시도_매칭키'].map(lambda x: KOR_COORDS.get(x, (None, None))[1])
    map_data = map_data.dropna(subset=['lat', 'lon'])

    col_map, col_top = st.columns([1.2, 1.8])
    
    with col_map:
        if not map_data.empty:
            fig_map = px.scatter_mapbox(
                map_data, lat="lat", lon="lon", size="용차단가(원)", color="용차단가(원)",
                hover_name="시도_매칭키", hover_data={"lat": False, "lon": False, "용차단가(원)": ":,.0f"},
                color_continuous_scale="Reds", size_max=40, zoom=5.5, center={"lat": 36.3, "lon": 127.8},
                mapbox_style="carto-positron", title="시/도별 출발 물동량 (표준 행정구역 기준)"
            )
            fig_map.update_layout(margin={"r":0,"t":40,"l":0,"b":0})
            st.plotly_chart(fig_map, use_container_width=True)
        else:
            st.info("💡 엑셀의 출발지가 '권역명'으로만 기재된 경우 지도에는 표시되지 않습니다. (우측 표에는 정상 집계됨)")
            
    with col_top:
        st.markdown(f"**💰 노선별 물동량 순위 &nbsp;&nbsp;|&nbsp;&nbsp; 총 {total_orders:,}건 (합계: {total_cost:,.0f}원)**")
        raw_df['노선'] = raw_df['출발지'] + " → " + raw_df['도착지']
        top_routes = raw_df.groupby('노선', as_index=False).agg(운송건수=('오더번호', 'count'), 용차단가_합계=('용차단가(원)', 'sum')).sort_values(by='용차단가_합계', ascending=False)
        
        display_top = top_routes.copy()
        display_top['용차단가_합계'] = display_top['용차단가_합계'].apply(lambda x: f"{x:,.0f} 원")
        display_top['운송건수'] = display_top['운송건수'].apply(lambda x: f"{x:,} 건")
        
        event = st.dataframe(display_top.head(10), use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
        
        if len(display_top) > 10:
            with st.expander("🔽 11위 이하 전체 구간 보기"):
                st.dataframe(display_top.iloc[10:], use_container_width=True, hide_index=True)
                
        if event.selection.rows:
            selected_route = top_routes.iloc[event.selection.rows[0]]['노선']
            st.markdown(f"**🔍 [{selected_route}] 상세 내역**")
            st.dataframe(raw_df[raw_df['노선'] == selected_route].drop(columns=['노선', '시도_매칭키']).sort_values('운송일자'), use_container_width=True, hide_index=True)

    st.divider()

    # --- 3. 왕복 배차 시뮬레이션 ---
    st.markdown("### 2️⃣ 왕복 배차 시뮬레이션 상세 결과 (물리적 매칭 & 용차 할인)")
    st.caption(f"자차를 투입하지 않고, 외부 용차를 왕복으로 묶었을 때 확보할 수 있는 {rt_discount_rate}%의 할인 효과 분석입니다.")
    
    if not rt_df.empty:
        rt_df['왕복배차_합계(원)'] = rt_df['용차비_합계(원)'] * (1 - rt_discount_rate / 100)
        rt_df['왕복배차_절감액(원)'] = rt_df['용차비_합계(원)'] - rt_df['왕복배차_합계(원)']
        
        rt_total_saving = rt_df['왕복배차_절감액(원)'].sum()
        rt_tobe_cost = total_cost - rt_total_saving
        
        kr1, kr2, kr3, kr4 = st.columns(4)
        kr1.metric("[AS-IS] 기존 편도 배차 총액", f"{total_cost:,.0f} 원", "전량 편도 사용 기준", delta_color="off")
        kr2.metric(f"[TO-BE 1] 왕복 할인({rt_discount_rate}%) 적용 시", f"{rt_tobe_cost:,.0f} 원", f"{rt_count}건 왕복 전환", delta_color="normal")
        kr3.metric("할인 기반 월간 절감액", f"{rt_total_saving:,.0f} 원", f"비용 {(rt_total_saving/total_cost)*100:.1f}% 절감", delta_color="normal")
        kr4.metric("적용 할인율 설정값", f"{rt_discount_rate} %", "사이드바에서 변경 가능", delta_color="off")
        
        st.markdown("#### 📈 왕복 배차 시 노선별 용차비 절감 효과 (Top 15)")
        rt_chart_df = rt_df.head(15).copy()
        rt_chart_df['시각화_라벨'] = rt_chart_df['구간'] + " (" + rt_chart_df['차종'] + " " + rt_chart_df['요구톤급(t)'].astype(str) + "t, " + rt_chart_df['상행_배차형태'] + ")"
        
        rt_long_df = pd.melt(rt_chart_df, id_vars=['시각화_라벨'], value_vars=['왕복배차_절감액(원)', '왕복배차_합계(원)'], 
                             var_name='비용 구분', value_name='금액')
        
        fig_rt_bar = px.bar(
            rt_long_df, x="금액", y="시각화_라벨", color="비용 구분", orientation='h', 
            color_discrete_map={'왕복배차_절감액(원)': px.colors.qualitative.Plotly[2], '왕복배차_합계(원)': px.colors.qualitative.Plotly[0]}, 
            title="막대 전체 길이: 개별 편도 용차 투입 시 발생 비용 합계", text_auto='.2s'
        )
        fig_rt_bar.update_layout(yaxis={'categoryorder':'total ascending'}, yaxis_title="노선 및 차량 스펙")
        st.plotly_chart(fig_rt_bar, use_container_width=True)

        st.markdown("#### 📋 왕복 배차 상세 테이블 (상/하행 분리)")
        display_rt_cols = ['운행시작일(상행)', '운행종료일(하행)', '구간', '차종', '요구톤급(t)', '상행_배차형태', '하행오더', '상행오더(리스트)', '상행금액(원)', '하행금액(원)', '용차비_합계(원)', '왕복배차_합계(원)', '왕복배차_절감액(원)']
        
        st.dataframe(
            rt_df[display_rt_cols].style.format({
                '요구톤급(t)': '{:.1f}', '상행금액(원)': '{:,.0f}', '하행금액(원)': '{:,.0f}', 
                '용차비_합계(원)': '{:,.0f}', '왕복배차_합계(원)': '{:,.0f}', '왕복배차_절감액(원)': '{:,.0f}'
            }), 
            use_container_width=True, height=300
        )
    else:
        st.info("조건을 만족하는 왕복 배차 조합이 없습니다.")
        
    st.divider()

    # --- 4. 자차 투입 시뮬레이션 ---
    st.markdown("### 3️⃣ 자차 투입 시뮬레이션 상세 결과 (Before vs After)")
    st.caption("위의 왕복 배차 물량들을 월간 단위로 묶어, 자차 고정비를 뺐을 때 흑자가 나는(효율 달성) 구간만 '투입 확정'으로 산출합니다.")
    
    if not summary_df.empty:
        rec_df = summary_df[summary_df['투입_판단'] == '자차 투입 (효율 달성)']
        
        own_fleet_cost = rec_df['월_자차운영비(B)'].sum()
        tobe_hired_cost = raw_df[~raw_df['오더번호'].isin(deploy_ids)]['용차단가(원)'].sum()
        tobe_total_cost = own_fleet_cost + tobe_hired_cost
        net_saving = total_cost - tobe_total_cost
        
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("[AS-IS] 기존 편도 배차 총액", f"{total_cost:,.0f} 원", "전량 용차 사용", delta_color="off")
        k2.metric("[TO-BE 2] 자차 투입 적용 시", f"{tobe_total_cost:,.0f} 원", f"{deploy_count}건 자차 흡수", delta_color="normal")
        k3.metric("자차 운영 기준 월간 절감액", f"{net_saving:,.0f} 원", f"비용 {(net_saving/total_cost)*100:.1f}% 절감", delta_color="normal")
        k4.metric("신규 필요 자차 대수", f"{rec_df['필요_자차대수'].sum():,.0f} 대", "추천 노선 합계", delta_color="off")
        
        st.markdown("#### 📈 자차 투입 시 노선/스펙별 기대 효과 (Top 15)")
        if not rec_df.empty:
            chart_df = rec_df.head(15).copy()
            chart_df['시각화_라벨'] = chart_df['운행_구간'] + " (" + chart_df['요구_차종'] + " " + chart_df['요구_톤급(t)'].astype(str) + "t, " + chart_df['상행_배차형태'] + ")"
            chart_df['잔여 운송비 (자차운영비)'] = chart_df['월_자차운영비(B)']
            
            long_df = pd.melt(chart_df, id_vars=['시각화_라벨'], value_vars=['월_비용절감액(A-B)', '잔여 운송비 (자차운영비)'], 
                              var_name='비용 구분', value_name='금액')
            
            fig_bar = px.bar(
                long_df, x="금액", y="시각화_라벨", color="비용 구분", orientation='h', 
                color_discrete_map={'월_비용절감액(A-B)': px.colors.qualitative.Plotly[2], '잔여 운송비 (자차운영비)': px.colors.qualitative.Plotly[0]}, 
                title="막대 전체 길이: 기존 용차 투입 시 발생 비용 (A)", text_auto='.2s'
            )
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, yaxis_title="노선 및 차량 스펙")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("자차 투입이 추천되는 노선이 없어 차트를 표시할 수 없습니다.")
        
        st.divider()

        st.markdown("#### 📋 노선/스펙별 자차 투입 분석 상세 (전체)")
        def color_status(val):
            if '효율 달성' in val: return 'background-color: #d4edda'
            elif '비효율' in val: return 'background-color: #fff3cd'
            else: return 'background-color: #f8d7da'
            
        styled_summary = summary_df.style.map(color_status, subset=['투입_판단']).format({
            '요구_톤급(t)': '{:.1f}', '대당_월고정비(원)': '{:,.0f}',
            '월_용차전환비(A)': '{:,.0f}', '월_자차운영비(B)': '{:,.0f}', '월_비용절감액(A-B)': '{:,.0f}'
        })
        
        event_summary = st.dataframe(
            styled_summary, use_container_width=True, hide_index=True,
            on_select="rerun", selection_mode="single-row"
        )
        
        if event_summary.selection.rows:
            sel_idx = event_summary.selection.rows[0]
            sel_route = summary_df.iloc[sel_idx]['운행_구간']
            sel_vtype = summary_df.iloc[sel_idx]['요구_차종']
            sel_ton = summary_df.iloc[sel_idx]['요구_톤급(t)']
            sel_dispatch = summary_df.iloc[sel_idx]['상행_배차형태']
            
            st.markdown(f"**🔍 [{sel_route} / {sel_vtype} {sel_ton}t / {sel_dispatch}] 월간 운행 상세 내역**")
            
            detail_rt = rt_df[(rt_df['구간'] == sel_route) & (rt_df['차종'] == sel_vtype) & (rt_df['요구톤급(t)'] == sel_ton) & (rt_df['상행_배차형태'] == sel_dispatch)]
            
            display_cols = ['운행시작일(상행)', '운행종료일(하행)', '하행오더', '상행오더(리스트)', '상행금액(원)', '하행금액(원)', '용차비_합계(원)']
            st.dataframe(
                detail_rt[display_cols].style.format({'상행금액(원)': '{:,.0f}', '하행금액(원)': '{:,.0f}', '용차비_합계(원)': '{:,.0f}'}), 
                use_container_width=True, hide_index=True
            )
        else:
            st.caption("☝️ 위 표의 행(Row)을 클릭하시면 해당 구간의 상세 운행(왕복) 내역을 확인할 수 있습니다.")
        
        out = BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='자차투입_월간분석')
            rt_df.to_excel(writer, index=False, sheet_name='왕복배차_할인내역')
            if not invalid_df.empty:
                invalid_df.to_excel(writer, index=False, sheet_name='제외된_불량데이터')
        st.download_button("📥 통합 분석 결과 엑셀 다운로드", data=out.getvalue(), file_name="최종_투입전략_분석결과.xlsx")
        
    else:
        st.info("분석 가능한 배차 조합이 없습니다.")

else:
    if not st.session_state.sim_run:
        st.info("👈 좌측 사이드바에서 엑셀 파일 2개를 업로드하고 [🚀 시뮬레이션 실행] 버튼을 눌러주세요.")