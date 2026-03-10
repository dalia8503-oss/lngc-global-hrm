  # 삭제 기능
        st.write("---")
        st.subheader("🗑️ 인력 삭제")
        col_del1, col_del2 = st.columns([3, 1])
        
        with col_del1:
            delete_emp_id = st.text_input("삭제할 사번 입력", max_chars=6, key="delete_emp_id")
        
        with col_del2:
            st.write("")  # 공간 맞추기
            if st.button("❌ 삭제하기", use_container_width=True):
                if delete_emp_id:
                    existing_worker = next((w for w in st.session_state.workers if w.get("사번") == delete_emp_id), None)
                    if existing_worker:
                        # 삭제 확인
                        st.session_state.confirm_delete_mode = True
                        st.session_state.pending_delete_emp_id = delete_emp_id
                        st.session_state.pending_delete_name = existing_worker.get('이름', '?')
                    else:
                        st.error(f"❌ 사번 {delete_emp_id}를 찾을 수 없습니다.")
                else:
                    st.error("사번을 입력해주세요.")
        
        # 삭제 확인 팝업
        if 'confirm_delete_mode' not in st.session_state:
            st.session_state.confirm_delete_mode = False
        
        if st.session_state.confirm_delete_mode:
            st.warning(f"⚠️ **사번 {st.session_state.pending_delete_emp_id} - {st.session_state.pending_delete_name} 정보를 삭제하시겠습니까?**")
            col_del_confirm1, col_del_confirm2 = st.columns(2)
            
            with col_del_confirm1:
                if st.button("✅ 네, 삭제하겠습니다", use_container_width=True, key="btn_confirm_delete"):
                    # 해당 사번의 인력 삭제
                    st.session_state.workers = [w for w in st.session_state.workers if w.get("사번") != st.session_state.pending_delete_emp_id]
                    
                    # 사진도 삭제
                    if st.session_state.pending_delete_emp_id in st.session_state.worker_photos:
                        del st.session_state.worker_photos[st.session_state.pending_delete_emp_id]
                    
                    save_data()
                    st.session_state.confirm_delete_mode = False
                    st.success(f"✅ 사번 {st.session_state.pending_delete_emp_id}가 삭제되었습니다.")
                    st.rerun()
            
            with col_del_confirm2:
                if st.button("❌ 아니오, 취소", use_container_width=True, key="btn_cancel_delete"):
                    st.session_state.confirm_delete_mode = False
                    st.rerun()






                            # 선택한 인력 정보 표시 및 수정 폼
        if st.session_state.edit_selected_emp_id:
            selected_worker = next((w for w in st.session_state.workers if w.get('사번') == st.session_state.edit_selected_emp_id), None)
            if selected_worker:
                st.session_state.selected_employee_data = selected_worker
                st.info(f"📋 선택된 인력: **{selected_worker.get('이름')}** (사번: {st.session_state.edit_selected_emp_id})")
                
                st.write("---")
                st.subheader("✏️ 정보 수정")
                
                # 부서 선택
                default_dept = selected_worker.get('부서') if selected_worker else None
                dept_list = list(dept_structure.keys())
                dept_index = dept_list.index(default_dept) if default_dept and default_dept in dept_list else 0
                edit_selected_dept = st.selectbox("소속 부서", dept_list, index=dept_index, key="edit_dept_select")
                    
                edit_col1, edit_col2 = st.columns(2)
                
                # 기본값 설정
                edit_default_emp_id = selected_worker.get('사번', '')
                edit_default_name = selected_worker.get('이름', '')
                edit_default_eng_name = selected_worker.get('영어이름', '')
                edit_default_nation = selected_worker.get('국적', '방글라데시')
                edit_default_entry = selected_worker.get('입국일', None)
                if edit_default_entry and isinstance(edit_default_entry, str):
                    edit_default_entry = datetime.strptime(edit_default_entry, '%Y-%m-%d').date()
                else:
                    edit_default_entry = datetime.now().date()
                
                with edit_col1:
                    edit_emp_id_input = st.text_input("사번", value=edit_default_emp_id, max_chars=6, key="edit_emp_id_input")
                    edit_name = st.text_input("이름", value=edit_default_name, key="edit_name")
                    edit_eng_name = st.text_input("영어 이름", value=edit_default_eng_name, key="edit_eng_name")
                    edit_nation = st.selectbox("국적", ["방글라데시", "파키스탄"], index=(0 if edit_default_nation == "방글라데시" else 1), key="edit_nation")
                    edit_photo = st.file_uploader("인물 사진 (선택 사항)", type=["jpg", "jpeg", "png"], key="edit_worker_photo")

                with edit_col2:
                    Unit_options = dept_structure[edit_selected_dept]["반"]
                    job_options = dept_structure[edit_selected_dept]["직종"]
                    
                    edit_default_unit = selected_worker.get('반') if selected_worker else None
                    edit_unit_index = Unit_options.index(edit_default_unit) if edit_default_unit and edit_default_unit in Unit_options else 0
                    edit_unit = st.selectbox("상세 조직(반)", Unit_options, index=edit_unit_index, key="edit_unit_select")
                    
                    edit_default_job = selected_worker.get('직종') if selected_worker else None
                    edit_job_index = job_options.index(edit_default_job) if edit_default_job and edit_default_job in job_options else 0
                    edit_job = st.selectbox("직종", job_options, index=edit_job_index, key="edit_job_select")
                    
                    edit_entry_date = st.date_input("입국일자", value=edit_default_entry, key="edit_entry_date")

                    # 입국일자부터 오늘까지의 개월수 계산 (소수점)
                    today = datetime.now().date()
                    edit_days_passed = (today - edit_entry_date).days
                    edit_calculated_months = round(edit_days_passed / 30.44, 1)
                    st.write(f"**근속개월**: {edit_calculated_months}개월")
                    edit_service_month = edit_calculated_months

                st.write("---")
                
                # 상세 정보 기본값 설정
                edit_default_h_type = selected_worker.get('숙소구분', '기숙사')
                edit_default_h_addr = selected_worker.get('주소', '')
                edit_default_c_type = selected_worker.get('계약', '부동산')
                edit_default_has_fam = selected_worker.get('가족동반', 'X')
                edit_default_fam_note = selected_worker.get('비고', '')
                
                with st.expander("📋 상세 정보"):
                    st.subheader("숙소 및 가족")
                    edit_col5, edit_col6 = st.columns(2)
                       
                    with edit_col5:
                        edit_h_type_index = 0 if edit_default_h_type == "기숙사" else 1
                        edit_h_type = st.radio("숙소 구분", ["기숙사", "사외"], index=edit_h_type_index, horizontal=True, key="edit_h_type")
                        edit_h_addr = st.text_input("숙소 상세 주소", value=edit_default_h_addr, key="edit_h_addr")

                    with edit_col6:
                        edit_c_type_index = 0 if edit_default_c_type == "부동산" else 1
                        edit_c_type = st.selectbox("계약 구분", ["부동산", "개인"], index=edit_c_type_index, key="edit_c_type")
                        edit_has_fam_index = 0 if edit_default_has_fam == "X" else 1
                        edit_has_fam = st.radio("가족 동반 여부", ["X", "O"], index=edit_has_fam_index, horizontal=True, key="edit_has_fam")
                           
                    edit_fam_note = st.text_area("비고 (가족 상세 및 기타 특이사항)", value=edit_default_fam_note, key="edit_fam_note")

                st.write("---")
                
                edit_col_btn1, edit_col_btn2 = st.columns(2)
                
                with edit_col_btn1:
                    if st.button("💾 수정 저장", use_container_width=True):
                        edit_updated_worker = {
                            "사번": edit_emp_id_input,
                            "이름": edit_name,
                            "영어이름": edit_eng_name,
                            "국적": edit_nation,
                            "부서": edit_selected_dept,
                            "반": edit_unit,
                            "직종": edit_job,
                            "근속개월": edit_service_month,
                            "입국일": edit_entry_date.strftime("%Y-%m-%d"),
                            "숙소구분": edit_h_type,
                            "주소": edit_h_addr,
                            "계약": edit_c_type,
                            "가족동반": edit_has_fam,
                            "비고": edit_fam_note
                        }
                        
                        # 사진 처리
                        if edit_photo:
                            try:
                                img = Image.open(edit_photo)
                                img_resized = img.resize((300, 500), Image.Resampling.LANCZOS)
                                st.session_state.worker_photos[edit_emp_id_input] = img_resized
                            except Exception as e:
                                st.warning(f"사진 처리 중 오류: {e}")
                        
                        # 기존 정보 업데이트
                        for i, worker in enumerate(st.session_state.workers):
                            if worker.get('사번') == st.session_state.edit_selected_emp_id:
                                st.session_state.workers[i] = edit_updated_worker
                                st.session_state.history.append({
                                    "사번": edit_emp_id_input,
                                    "변경일": datetime.now().strftime("%Y-%m-%d"),
                                    "내용": f"{edit_selected_dept} {edit_unit} 정보 수정"
                                })
                                break
                        
                        save_data()
                        st.session_state.selected_employee_data = None
                        st.session_state.edit_selected_emp_id = None
                        st.success(f"✅ 사번 {edit_emp_id_input} 정보가 수정되었습니다.")
                        st.rerun()
                
                with edit_col_btn2:
                    if st.button("❌ 취소", use_container_width=True):
                        st.session_state.selected_employee_data = None
                        st.session_state.edit_selected_emp_id = None
                        st.rerun()