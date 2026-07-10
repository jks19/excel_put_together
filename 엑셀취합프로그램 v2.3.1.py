import xlwings as xw
import os
import shutil
import sys
import re
import pickle
from collections import defaultdict
import pandas as pd

class ExcelConsolidator:
    def __init__(self, app):
        if getattr(sys, 'frozen', False):
            # 패키징된 exe 실행 환경
            base_path = os.path.dirname(sys.executable)
        else:
            # 일반 파이썬 스크립트 실행 환경
            base_path = os.path.dirname(os.path.abspath(__file__))

        self.app = app

        # 기본 경로
        self.base_path = base_path
        self.template_path = os.path.join(base_path, "양식")
        self.input_folder = os.path.join(base_path, "취합")
        self.output_folder = os.path.join(base_path, "결과")
        self.state_file = os.path.join(base_path, "결과", "consolidation_state.pkl")

        # 추가 생성 가능 폴더 경로
        self.processed_folder = os.path.join(self.input_folder, "_처리완료")
        self.conflict_folder = os.path.join(self.input_folder, "_오류", "충돌")
        self.error_subfolder = os.path.join(self.input_folder, "_오류", "처리오류")
        self.error_folder = os.path.join(self.input_folder, "_오류")
        # 공통 옵션 / # 
        self.file_type = ('.xlsx', '.xls', '.xlsm')
        self.blue_color = (0, 176, 240)
        self.changed_cells = defaultdict(dict)  # {sheet_name: {coord: {'filename': str, 'value': any}}}
        self.conflict_files = []
        self.error_files = []
        self.processed_files = []

    def load_state(self):
        """이전 취합 상태 로드"""
        try:
            with open(self.state_file, 'rb') as f:
                self.changed_cells = pickle.load(f)
            print(f"✓ 이전 상태 로드됨: {len(self.changed_cells)} 시트\n")
        except Exception as e:
            print(f"⚠️  상태 파일 로드 실패: {e}\n")

    def save_state(self):
        """현재 취합 상태 저장"""
        try:
            with open(self.state_file, 'wb') as f:
                pickle.dump(self.changed_cells, f)
            print(f"\n✓ 상태 저장 완료: {self.state_file}")
        except Exception as e:
            print(f"⚠️  상태 저장 실패: {e}")

    def create_directory_structure(self):
        """필요한 폴더 구조 생성 (오류 폴더 제외)"""
        os.makedirs(self.template_path, exist_ok=True)
        os.makedirs(self.input_folder, exist_ok=True)
        os.makedirs(self.output_folder, exist_ok=True)
        os.makedirs(os.path.join(self.input_folder, "_처리완료"), exist_ok=True)
        
        print("📁 작업 폴더 구조:")
        print(f"  양식: {self.template_path}")
        print(f"  취합: {self.input_folder}")
        print(f"  결과: {self.output_folder}\n")

    def create_conflict_folders(self):
        """충돌 폴더 구조 생성 (충돌 발생 시에만)"""
        os.makedirs(os.path.join(self.input_folder, "_오류", "충돌"), exist_ok=True)

    def create_error_subfolders(self):
        """오류 폴더 구조 생성 (처리오류 발생 시에만)"""
        os.makedirs(os.path.join(self.input_folder, "_오류", "처리오류"), exist_ok=True)
    
    def check_template_file(self):
        """템플릿 파일 확인 (while로 재귀 처리)"""
        while True:
            template_files = [
                f for f in os.listdir(self.template_path)
                if f.endswith(self.file_type) and not f.startswith('~')
            ]
            
            if not template_files:
                print("❌ 오류: '양식' 폴더에 양식 파일이 없습니다.")
                print(f"📍 경로: {self.template_path}")
                print("    양식 파일(*.xlsx 또는 *.xls)을 위 폴더에 넣어주세요.\n")
                
                input("파일을 추가한 후 엔터를 눌러주세요: ")
                continue
            
            if len(template_files) == 1:
                return os.path.join(self.template_path, template_files[0])
            
            # 2개 이상인 경우
            print(f"⚠️  경고: '양식' 폴더에 {len(template_files)}개의 파일이 있습니다.")
            print("    양식 파일은 1개만 있어야 합니다.\n")
            for i, f in enumerate(template_files, 1):
                print(f"  {i}. {f}")
            
            print("\n불필요한 파일을 삭제하고 1개만 남겨주세요.")
            input("정리한 후 엔터를 눌러주세요: ")
    
    def check_output_files(self):
        """결과 파일 확인 (while로 재귀 처리)"""
        while True:
            output_files = [
                f for f in os.listdir(self.output_folder)
                if f.endswith(self.file_type) and not f.startswith('~')
            ]
            
            if not output_files:
                # 최초 취합: 새로운 결과 파일 생성
                return os.path.join(self.output_folder, "취합결과.xlsx")
            
            if len(output_files) == 1:
                # 결과 파일 1개: 이전 상태에서 계속
                result_file = os.path.join(self.output_folder, output_files[0])
                print(f"✓ 기존 결과 파일을 감지했습니다.")
                print(f"  파일: {output_files[0]}")
                print(f"  기존 취합 결과에 이어서 진행합니다.\n")
                return result_file
            
            # 2개 이상: 사용자에게 정리 요청
            print(f"⚠️  경고: '결과' 폴더에 {len(output_files)}개의 파일이 있습니다.")
            print("    파일은 1개만 있어야 합니다.\n")
            for i, f in enumerate(output_files, 1):
                print(f"  {i}. {f}")
            
            print("\n다음 중 하나를 선택해주세요:")
            print("  1. 이어서 취합할 파일만 남기고 나머지 삭제")
            print("  2. 모든 파일을 삭제하고 새로 시작\n")
            
            input("위 작업을 완료한 후 엔터를 눌러주세요: ")
    
    def check_input_files(self):
        """입력 폴더 파일 확인 (while 재귀)"""
        while True:
            input_files = [
                f for f in os.listdir(self.input_folder)
                if f.endswith(self.file_type) and not f.startswith('~')
            ]
            
            if input_files:
                return sorted(input_files)
            
            # 파일 없음
            print("⚠️  '취합' 폴더에 처리할 파일이 없습니다.")
            print(f"📍 경로: {self.input_folder}\n")
            print("처리할 파일들을 '취합' 폴더에 넣어주세요.")
            
            input("\n파일을 추가한 후 엔터를 눌러주세요: ")

    def get_all_coords(self, ws1, ws2):
        """두 시트의 최대 행/열을 기준으로 모든 셀 좌표 반환"""
        max_row = max(ws1.used_range.last_cell.row, ws2.used_range.last_cell.row)
        max_col = max(ws1.used_range.last_cell.column, ws2.used_range.last_cell.column)
        
        all_coords = set()
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = ws1.cells(row, col)
                all_coords.add(cell.address)
        
        return all_coords
    
    def get_cell_value(self, ws, address):
        """셀 값을 안전하게 가져오기"""
        try:
            return ws.range(address).value
        except:
            return None
    
    def set_cell_value(self, ws, address, value):
        """셀 값 설정"""
        try:
            ws.range(address).value = value
        except Exception as e:
            print(f"셀 값 설정 실패 {address}: {e}")
    
    def set_cell_color(self, ws, address, rgb_color):
        """셀 색상 설정"""
        try:
            ws.range(address).color = rgb_color
        except Exception as e:
            print(f"셀 색상 설정 실패 {address}: {e}")

    def is_formula(self, ws, address):
        """셀이 수식인지 확인"""
        try:
            cell = ws.range(address)
            return cell.formula.startswith('=')
        except:
            return False

    def compare_worksheets(self, template_ws, source_ws):
        """두 시트를 범위 단위로 읽어 비교하고 변경된 셀 반환"""
        max_row = max(
            template_ws.used_range.last_cell.row,
            source_ws.used_range.last_cell.row
        )
        max_col = max(
            template_ws.used_range.last_cell.column,
            source_ws.used_range.last_cell.column
        )

        template_range = template_ws.range((1, 1), (max_row, max_col))
        source_range = source_ws.range((1, 1), (max_row, max_col))

        # Excel COM 호출을 최소화하기 위해 범위를 한 번에 읽는다.
        template_values = template_range.options(ndim=2).value
        source_values = source_range.options(ndim=2).value
        template_formulas = template_range.formula

        changes = {}

        for row_idx in range(max_row):
            for col_idx in range(max_col):
                formula = template_formulas[row_idx][col_idx]

                # 기존 기능과 동일하게 템플릿의 수식 셀은 비교에서 제외
                if isinstance(formula, str) and formula.startswith('='):
                    continue

                template_value = template_values[row_idx][col_idx]
                source_value = source_values[row_idx][col_idx]

                if template_value != source_value:
                    coord = f"{xw.utils.col_name(col_idx + 1)}{row_idx + 1}"
                    changes[coord] = source_value

        return changes
    
    def apply_changes_to_template(self, result_ws, changes):
        """템플릿에 변경사항 적용"""
        for coord, value in changes.items():
            self.set_cell_value(result_ws, coord, value)
            self.set_cell_color(result_ws, coord, self.blue_color)
    
    def has_conflict(self, sheet_name, changes):
        """충돌 여부 확인: 이미 변경된 셀 중복 체크"""
        if sheet_name not in self.changed_cells:
            return None
        
        for coord in changes:
            if coord in self.changed_cells[sheet_name]:
                return coord
        return None
    
    def record_changes(self, sheet_name, changes, filename):
        """변경된 셀 기록 (파일명 함께 저장)
        
        예시:
        - sheet_name: "Sheet1"
        - changes: {'A1': value1, 'B2': value2}
        - filename: "답변_01.xlsx"
        - self.changed_cells["Sheet1"] = {
            'A1': {'filename': '답변_01.xlsx', 'value': value1},
            'B2': {'filename': '답변_01.xlsx', 'value': value2}
          }
        """
        for coord, value in changes.items():
            self.changed_cells[sheet_name][coord] = {
                'filename': filename,
                'value': value
            }
    
    def input_excel_cell(self):
        pattern = r'^[A-Za-z]+[1-9][0-9]*$'
        while True:
            user_input = input('첫번째 칼럼 셀 위치를 입력하세요.(예: A4): ')
            if re.match(pattern, user_input):
                return user_input
            else:
                print('잘못된 형식입니다. 예) A4, B12 형식으로 입력해 주세요.')

    def append_to_template_position(self):
        """모든 파일 취합 시작"""
        # 폴더 생성
        self.create_directory_structure()
        
        # 템플릿 확인
        template_file = self.check_template_file()
        
        # 결과 파일 확인 (경로 반환, 없으면 새 경로)
        result_file = self.check_output_files()
        
        # 입력 파일 확인
        input_files = self.check_input_files()
        
        try:
            template_wb = self.app.books.open(template_file)

        except Exception as e:
            print(f"❌ 양식 파일 열기 실패: {e}")
            return
        
        # 결과 파일 생성/로드
        if os.path.exists(result_file) and os.path.exists(self.state_file):
            # 기존 파일: 상태 복원
            try:
                result_wb = self.app.books.open(result_file)
                self.load_state()
            except Exception as e:
                print(f"❌ 기존 결과 파일 열기 실패: {e}")
                return
        else:
            # 새 파일: 템플릿 복사
            shutil.copy(template_file, result_file)
            try:
                result_wb = self.app.books.open(result_file)
            except Exception as e:
                print(f"❌ 결과 파일 생성 실패: {e}")
                return

        # 결과 파일에 시트 및 통합문서 보호 설정 해제
        wb_pw = input('통합문서 보호 암호를 입력하세요. 없으면 엔터를 누르세요.')
        ws_pw = input('워크시트 보호 암호를 입력하세요. 없으면 엔터를 누르세요.')
        result_wb.api.Unprotect(Password=f'{wb_pw}')

        template_sheet_names = [sheet.name for sheet in template_wb.sheets]

        for sheet_name in template_sheet_names:
            result_ws = result_wb.sheets[sheet_name]
            result_ws.api.Unprotect(Password=f'{ws_pw}')

        # 입력 파일 가져오기
        print(f"총 {len(input_files)}개 파일 처리 시작...")
        
        processed_count = 0
        error_count = 0
        error_msgs = []

        for idx, filename in enumerate(input_files, 1):
            # 취합하려는 파일명이 양식 파일과 동일한 경우 자동으로 취합파일명 수정하고 진행하기(맨 뒤에 _ 붙여서)
            if filename == template_file.split('\\')[-1]:
                name, ext = os.path.splitext(filename)  # 이름과 확장자 분리
                os.rename(os.path.join(self.input_folder, filename), os.path.join(self.input_folder, f'{name}_임시복제본{ext}'))
                filename = f'{name}_임시복제본{ext}'
            
            file_path = os.path.join(self.input_folder, filename)

            try:
                current_wb = self.app.books.open(file_path)
                
                # 1단계: 모든 시트 검증 및 변경사항 추출
                changes_by_sheet = {}
                file_has_error = False
                error_sheet = None
                error_coord = None
                error_origin_file = None

                current_sheet_names = [sheet.name for sheet in current_wb.sheets]
                # if set(template_sheet_names) != set(current_sheet_names):
                diff1 = set(template_sheet_names) - set(current_sheet_names)    # 임의로 답변받아야 할 시트를 제거한 답변파일이 있는 경우
                if diff1:
                    file_has_error = True
                    error_sheet = diff1
                    # diff1 = set(template_sheet_names) - set(current_sheet_names)    # 임의로 답변받아야 할 시트를 제거한 답변파일이 있는 경우
                    # diff2 = set(current_sheet_names) - set(template_sheet_names)    # 임의로 시트를 추가한 답변파일이 있는 경우 / template 파일에서 일부 시트를 지운 경우(현재 임시로 정상)
                    # error_sheet = diff1 | diff2
                else:
                    for sheet_name in template_sheet_names:     # $ 잠기지 않은 셀이 있는 시트만 하면 더 좋을 듯 // 일단은 template 파일에서 취합할 시트만 남겨서 진행하는 방식으로 사용
                        try:
                            if sheet_name not in current_sheet_names:
                                file_has_error = True
                                error_sheet = sheet_name
                                break

                            template_ws = template_wb.sheets[sheet_name]
                            current_ws = current_wb.sheets[sheet_name]
                            result_ws = result_wb.sheets[sheet_name]
                            
                            changes = self.compare_worksheets(template_ws, current_ws)
                            
                            if changes:
                                conflict_coord = self.has_conflict(sheet_name, changes)
                                if conflict_coord:
                                    file_has_error = True
                                    error_sheet = sheet_name
                                    error_coord = conflict_coord
                                    error_origin_file = self.changed_cells[sheet_name][error_coord]['filename']
                                    break
                                
                                changes_by_sheet[sheet_name] = changes
                        
                        except KeyError:
                            err_msg = f"[ERROR] {filename} - 시트 '{sheet_name}' 없음"
                            print(err_msg)
                            error_msgs.append(err_msg)
                            file_has_error = True
                            error_sheet = sheet_name
                            break
                        except Exception as e:
                            err_msg = f"[ERROR] {filename} 처리 중 오류: {str(e)}"
                            print(err_msg)
                            error_msgs.append(err_msg)
                            file_has_error = True
                            break
                    
                current_wb.close()
                
                # 2단계: 에러 있으면 파일만 이동
                if file_has_error:
                    err_msg = f"\n⚠️  [충돌/오류 감지] {filename}"
                    print(err_msg)
                    error_msgs.append(err_msg)
                    if error_sheet:
                        self.create_conflict_folders()
                        shutil.move(file_path, os.path.join(self.conflict_folder, filename))
                        self.conflict_files.append(filename)

                        if error_coord:
                            err_msg = f"   시트: {error_sheet}, 충돌 셀: {error_coord}, 충돌 파일: {error_origin_file}"
                            print(err_msg)
                            error_msgs.append(err_msg)
                            shutil.move(os.path.join(self.processed_folder, error_origin_file), os.path.join(self.conflict_folder, error_origin_file))
                            # 원본 파일에서 변경됐던 셀들을 template 상태로 되돌리기
                            print(f"   원본 파일의 변경사항을 되돌리고 있습니다...")
                            for sheet_name_key in self.changed_cells:
                                coords_to_revert = []
                                for coord, info in self.changed_cells[sheet_name_key].items():
                                    if info['filename'] == error_origin_file:
                                        coords_to_revert.append(coord)
                                
                                if coords_to_revert:
                                    result_ws = result_wb.sheets[sheet_name_key]
                                    template_ws = template_wb.sheets[sheet_name_key]
                                    
                                    for coord in coords_to_revert:
                                        template_value = self.get_cell_value(template_ws, coord)
                                        self.set_cell_value(result_ws, coord, template_value)
                                        # 파란색 제거 (template 상태로 원상복구)
                                        self.set_cell_color(result_ws, coord, template_ws.range(coord).color)
                                        
                                        # changed_cells에서도 제거
                                        del self.changed_cells[sheet_name_key][coord]
                            processed_count -= 1

                            print(f"   ✓ 원본 파일의 변경사항 복원 완료")
                        else:
                            err_msg = f"   시트: {error_sheet}"
                            print(err_msg)
                            error_msgs.append(err_msg)
                    else:
                        self.create_error_subfolders()
                        shutil.move(file_path, os.path.join(self.error_subfolder, filename))
                        self.error_files.append(filename)

                    print("   → 파일 제외\n")
                    error_count += 1
                else:
                    # 3단계: 에러 없으면 모든 변경사항 적용
                    for sheet_name, changes in changes_by_sheet.items():
                        result_ws = result_wb.sheets[sheet_name]
                        self.apply_changes_to_template(result_ws, changes)
                        self.record_changes(sheet_name, changes, filename)
                    
                    processed_file_path = os.path.join(self.processed_folder, filename)
                    shutil.move(file_path, processed_file_path)
                    self.processed_files.append(filename)
                    print(f"[{idx}/{len(input_files)}] {filename} - 처리 완료 ✓")
                    processed_count += 1
                
            except Exception as e:
                err_msg = f"[ERROR] {filename} 처리 중 심각한 오류: {str(e)}\n"
                print(err_msg)
                print("   → 파일 제외\n")
                error_msgs.append(err_msg)
                self.create_error_subfolders()
                shutil.move(file_path, os.path.join(self.error_subfolder, filename))
                self.error_files.append(filename)
                error_count += 1
        
        # 저장 및 닫기
        try:
            result_wb.save()
            result_wb.close()
            template_wb.close()
        except Exception as e:
            print(f"파일 저장 중 오류: {e}")
            input('종료하려면 아무키나 누르세요.')

        # 상태 저장
        self.save_state()
        
        # 완료 보고
        print("\n" + "="*60)
        print(f"취합 완료!")
        print(f"처리된 파일: {processed_count}개")
        # print(f"오류 파일: {error_count}개")
        print(f"\n📄 결과 파일: {result_file}")
        print("="*60)
        
        # 에러 폴더 열기 (1건 이상)
        if error_count > 0:
            print(f"\n❌ 오류 발생 파일 ({error_count}개) 내역 요약")
            print(f'{'\n'.join(error_msgs)}')
            print(f"📁 오류 파일을 확인하고 수정하여 '취합' 폴더에 다시 넣고 재실행하세요.")
            os.startfile(self.error_folder)
        else:
            # 성공 시 결과 파일 열기
            print(f"\n✅ 모든 파일이 안전하게 처리되었습니다!")
            print(f"\n📄 결과 파일을 열고 있습니다...\n")
            os.startfile(os.path.dirname(result_file))

    def concat_files(self):     # $$ 미확인
        """모든 파일 취합 시작"""
        # 폴더 생성
        self.create_directory_structure()
        
        # 템플릿 확인
        template_file = self.check_template_file()

        # 결과 파일 확인 (경로 반환, 없으면 새 경로)
        result_file = self.check_output_files()
        
        # 입력 파일 확인
        input_files = self.check_input_files()

        start_cell = self.input_excel_cell()

        header_row = int(re.search(r'\\d+', start_cell).group())
        try:
            template_wb = self.app.books.open(template_file)
            template_sheet_names = [sheet.name for sheet in template_wb.sheets]
            template_wb.close()
            template_cols = {}
            for sheet_name in template_sheet_names:
                template_df = pd.read_excel(result_file, header=header_row, sheet_name=sheet_name)
                template_cols[sheet_name] = set(template_df.columns)
                offset_cnt = len(template_df)
            del template_df

        except Exception as e:
            print(f"❌ 양식 파일 열기 실패: {e}")
            return

        # 새 파일: 템플릿 복사
        shutil.copy(template_file, result_file)
        try:
            result_wb = self.app.books.open(result_file)
        except Exception as e:
            print(f"❌ 결과 파일 생성 실패: {e}")
            return
        
        
        
        # 입력 파일 가져오기
        print(f"총 {len(input_files)}개 파일 처리 시작...")
        
        processed_count = 0
        error_count = 0
        error_msgs = []


        result_dfs = defaultdict(list)
        changes_list = []
        for idx, filename in enumerate(input_files, 1):
            file_path = os.path.join(self.input_folder, filename)

            try:
                file_has_error = False

                changes = {}
                for sheet_name in template_sheet_names:
                    current_ws = pd.read_excel(file_path, header=header_row, sheet_name=sheet_name)
                    if template_cols[sheet_name] != set(current_ws.columns):
                        ## 에러 처리
                        error_sheet_name = sheet_name
                        file_has_error = True
                        break
                    else:
                        current_ws.loc[:, '출처 파일명'] = filename
                        changes[sheet_name] = current_ws

                if file_has_error:
                    self.create_error_subfolders()
                    shutil.move(file_path, os.path.join(self.error_subfolder, filename))
                    self.error_files.append(filename)
                    
                    err_msg = f"\n⚠️  [충돌/오류 감지] {filename}\n   칼럼명이 불일치 합니다.\n   시트: {error_sheet_name}, 칼럼: {list(current_ws.columns)}"
                    error_msgs.append(err_msg)
                    print("   → 파일 제외\n")
                    error_count += 1

                else:
                    changes_list.append(changes)
                    processed_file_path = os.path.join(self.processed_folder, filename)
                    shutil.move(file_path, processed_file_path)
                    self.processed_files.append(filename)
                    print(f"[{idx}/{len(input_files)}] {filename} - 처리 완료 ✓")
                    processed_count += 1

            except Exception as e:
                err_msg = f"[ERROR] {filename} 처리 중 심각한 오류: {str(e)}\n"
                print(err_msg)
                print("   → 파일 제외\n")
                error_msgs.append(err_msg)
                self.create_error_subfolders()
                shutil.move(file_path, os.path.join(self.error_subfolder, filename))
                self.error_files.append(filename)
                error_count += 1

        # 저장 및 닫기
        try:
            # 취합을 위해 result_dfs(dict)에 저장
            for d in changes_list:
                for k, v in d.items():
                    v = v.iloc[offset_cnt:, :]   # 양식의 예시 행 있다면 제거
                    result_dfs[k].append(v)

            # result_dfs(dict) 순회하면서 concat하여 저장하기
            for sheet_name in template_sheet_names:
                result_ws = result_wb.sheets[sheet_name]

                start_cell = result_ws.range(start_cell)
                insert_cell = start_cell.offset(1, 0)   # 바로 아래 셀 위치로 이동 (행: +1, 열: 0)

                ws_concat_df = pd.concat(result_dfs[sheet_name], axis=0)    # DataFrame 값만 삽입 (헤더 없이)
                ws_concat_df = ws_concat_df.drop_duplicates()
                insert_cell.options(index=False, header=False).value = ws_concat_df

            result_wb.save()
            result_wb.close()

        except Exception as e:
            print(f"파일 저장 중 오류: {e}")
            input('종료하려면 아무키나 누르세요.')

        # 완료 보고
        print("\n" + "="*60)
        print(f"취합 완료!")
        print(f"처리된 파일: {processed_count}개")
        # print(f"오류 파일: {error_count}개")
        print(f"\n📄 결과 파일: {result_file}")
        print("="*60)
        
        # 에러 폴더 열기 (1건 이상)
        if error_count > 0:
            print(f"\n❌ 오류 발생 파일 ({error_count}개) 내역 요약")
            print(f'{'\n'.join(error_msgs)}')
            print(f"📁 오류 파일을 확인하고 수정하여 '취합' 폴더에 다시 넣고 재실행하세요.")
            os.startfile(self.error_folder)
        else:
            # 성공 시 결과 파일 열기
            print(f"\n✅ 모든 파일이 안전하게 처리되었습니다!")
            print(f"\n📄 결과 파일 폴더를 열고 있습니다...\n")
            os.startfile(os.path.dirname(result_file))


# 사용 예제
if __name__ == "__main__":

    # app = xw.App(visible=True)     # 작업용: 엑셀 창 실시간으로 보면서 확인 가능
    app = xw.App(visible=False)

    try:
        consolidator = ExcelConsolidator(app)
        consolidator.append_to_template_position()
        input('종료하려면 아무키나 누르세요.')
    finally:
        app.quit()