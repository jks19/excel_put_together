import xlwings as xw
import os
import shutil
import sys

class ExcelConsolidator:
    def __init__(self):
        if getattr(sys, 'frozen', False):
            # 패키징된 exe 실행 환경
            base_path = os.path.dirname(sys.executable)
        else:
            # 일반 파이썬 스크립트 실행 환경
            base_path = os.path.dirname(os.path.abspath(__file__))


        # 기본 경로
        self.base_path = base_path
        self.template_path = os.path.join(base_path, "template")
        self.input_folder = os.path.join(base_path, "input")
        self.output_folder = os.path.join(base_path, "output")
        
        # 추가 생성 가능 폴더 경로
        self.processed_folder = os.path.join(self.input_folder, "_처리완료")
        self.conflict_folder = os.path.join(self.input_folder, "_오류", "충돌")
        self.error_subfolder = os.path.join(self.input_folder, "_오류", "처리오류")
        self.error_folder = os.path.join(self.input_folder, "_오류")

        self.blue_color = (0, 112, 192)
        self.changed_cells = {}  # {sheet_name: set(coords)}
        self.conflict_files = []
        self.error_files = []
        self.processed_files = []
        
    def create_directory_structure(self):
        """필요한 폴더 구조 생성 (오류 폴더 제외)"""
        os.makedirs(self.template_path, exist_ok=True)
        os.makedirs(self.input_folder, exist_ok=True)
        os.makedirs(self.output_folder, exist_ok=True)
        os.makedirs(os.path.join(self.input_folder, "_처리완료"), exist_ok=True)
        
        print("📁 작업 폴더 구조:")
        print(f"  template: {self.template_path}")
        print(f"  input: {self.input_folder}")
        print(f"  output: {self.output_folder}\n")

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
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')
            ]
            
            if not template_files:
                print("❌ 오류: template 폴더에 양식 파일이 없습니다.")
                print(f"📍 경로: {self.template_path}")
                print("    양식 파일(*.xlsx 또는 *.xls)을 위 폴더에 넣어주세요.\n")
                
                input("파일을 추가한 후 엔터를 눌러주세요: ")
                continue
            
            if len(template_files) == 1:
                return os.path.join(self.template_path, template_files[0])
            
            # 2개 이상인 경우
            print(f"⚠️  경고: template 폴더에 {len(template_files)}개의 파일이 있습니다.")
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
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')
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
            print(f"⚠️  경고: output 폴더에 {len(output_files)}개의 파일이 있습니다.")
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
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')
            ]
            
            if input_files:
                return sorted(input_files)
            
            # 파일 없음
            print("⚠️  input 폴더에 처리할 파일이 없습니다.")
            print(f"📍 경로: {self.input_folder}\n")
            print("처리할 파일들을 input 폴더에 넣어주세요.")
            
            input("\n파일을 추가한 후 엔터를 눌러주세요: ")
    
    def build_changed_cells_from_result(self, template_wb, result_wb):
        """
        기존 결과 파일과 템플릿을 비교하여 changed_cells 구축
        """
        for sheet_name in [sheet.name for sheet in template_wb.sheets]:
            try:
                template_ws = template_wb.sheets[sheet_name]
                result_ws = result_wb.sheets[sheet_name]
                
                all_coords = self.get_all_coords(template_ws, result_ws)
                
                for coord in all_coords:
                    # 수식인 경우 제외 $ 잠금해제된 것만 하면 더 좋을 듯
                    if self.is_formula(template_ws, coord):
                        continue

                    template_value = self.get_cell_value(template_ws, coord)
                    result_value = self.get_cell_value(result_ws, coord)
                    
                    # 결과 파일이 템플릿과 다르면 변경된 것
                    if template_value != result_value:
                        if sheet_name not in self.changed_cells:
                            self.changed_cells[sheet_name] = set()
                        self.changed_cells[sheet_name].add(coord)
            except Exception as e:
                print(f"❌ 오류: 결과 파일 상태 복원 실패")
                print(f"   시트: {sheet_name}")
                print(f"   오류: {str(e)}")
                print("\n프로그램을 종료합니다.")
                sys.exit(1)

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
        """두 시트를 비교하고 변경된 셀 반환"""
        all_coords = self.get_all_coords(template_ws, source_ws)
        
        changes = {}
        for coord in all_coords:
            # 수식인 경우 제외 $ 잠금해제된 것만 하면 더 좋을 듯
            if self.is_formula(template_ws, coord):
                continue

            template_value = self.get_cell_value(template_ws, coord)
            source_value = self.get_cell_value(source_ws, coord)
            
            if template_value != source_value:
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
    
    def record_changes(self, sheet_name, changes):
        """변경된 셀 기록
        
        예시:
        - sheet_name: "Sheet1"
        - changes.keys(): dict_keys(['A1', 'B2', 'C3'])
        - self.changed_cells["Sheet1"] = {'A1', 'B2', 'C3'}
        - update() 후: {'A1', 'B2', 'C3', 'D4'} (새 요소 추가)
        """
        if sheet_name not in self.changed_cells:
            self.changed_cells[sheet_name] = set()
        
        self.changed_cells[sheet_name].update(changes.keys())
    
    def open_folder(self, folder_path):
        """폴더 열기"""
        try:
            if sys.platform == 'win32':
                os.startfile(folder_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{folder_path}"')
            else:
                os.system(f'xdg-open "{folder_path}"')
        except Exception as e:
            print(f"폴더 열기 실패: {e}")
    
    def consolidate(self):
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
            template_wb = xw.Book(template_file, visible=False)
        except Exception as e:
            print(f"❌ 템플릿 파일 열기 실패: {e}")
            return
        
        # 결과 파일 생성/로드
        if os.path.exists(result_file):
            # 기존 파일: 상태 복원
            try:
                result_wb = xw.Book(result_file)
                self.build_changed_cells_from_result(template_wb, result_wb)
            except Exception as e:
                print(f"❌ 결과 파일 열기 실패: {e}")
                return
        else:
            # 새 파일: 템플릿 복사
            shutil.copy(template_file, result_file)
            try:
                result_wb = xw.Book(result_file)
            except Exception as e:
                print(f"❌ 결과 파일 생성 실패: {e}")
                return
        
        template_sheet_names = [sheet.name for sheet in template_wb.sheets]
        
        # 입력 파일 가져오기
        print(f"총 {len(input_files)}개 파일 처리 시작...")
        
        processed_count = 0
        error_count = 0
        
        for idx, filename in enumerate(input_files, 1):
            file_path = os.path.join(self.input_folder, filename)
            
            try:
                current_wb = xw.Book(file_path)
                
                # 1단계: 모든 시트 검증 및 변경사항 추출
                changes_by_sheet = {}
                file_has_error = False
                error_sheet = None
                error_coord = None

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
                                    break
                                
                                changes_by_sheet[sheet_name] = changes
                        
                        except KeyError:
                            print(f"[ERROR] {filename} - 시트 '{sheet_name}' 없음")
                            file_has_error = True
                            error_sheet = sheet_name
                            break
                        except Exception as e:
                            print(f"[ERROR] {filename} 처리 중 오류: {str(e)}")
                            file_has_error = True
                            break
                    
                current_wb.close()
                
                # 2단계: 에러 있으면 파일만 이동
                if file_has_error:
                    print(f"\n⚠️  [충돌/오류 감지] {filename}")
                    if error_sheet:
                        self.create_conflict_folders()
                        shutil.move(file_path, os.path.join(self.conflict_folder, filename))
                        self.conflict_files.append(filename)

                        if error_coord:
                            print(f"   시트: {error_sheet}, 충돌 셀: {error_coord}")
                        else:
                            print(f"   시트: {error_sheet}")
                    else:
                        self.create_error_subfolders()
                        shutil.move(file_path, os.path.join(self.error_subfolder, filename))
                        self.error_files.append(filename)

                    print(f"   → 파일 제외\n")
                    error_count += 1
                else:
                    # 3단계: 에러 없으면 모든 변경사항 적용
                    for sheet_name, changes in changes_by_sheet.items():
                        result_ws = result_wb.sheets[sheet_name]
                        self.apply_changes_to_template(result_ws, changes)
                        self.record_changes(sheet_name, changes)
                    
                    processed_file_path = os.path.join(self.processed_folder, filename)
                    shutil.move(file_path, processed_file_path)
                    self.processed_files.append(filename)
                    print(f"[{idx}/{len(input_files)}] {filename} - 처리 완료 ✓")
                    processed_count += 1
                
            except Exception as e:
                print(f"[ERROR] {filename} 처리 중 심각한 오류: {str(e)}")
                self.create_error_folders()
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
        
        # 완료 보고
        print("\n" + "="*60)
        print(f"취합 완료!")
        print(f"처리된 파일: {processed_count}개")
        # print(f"오류 파일: {error_count}개")
        print(f"\n📄 결과 파일: {result_file}")
        print("="*60)
        
        # 에러 폴더 열기 (1건 이상)
        if error_count > 0:
            print(f"\n❌ 오류 발생 파일 ({error_count}개)")
            print(f"📁 오류 파일을 확인하세요.")
            self.open_folder(self.error_folder)
        else:
            # 성공 시 결과 파일 열기
            print(f"\n✅ 모든 파일이 안전하게 처리되었습니다!")
            print(f"\n📁 결과 파일을 열고 있습니다...\n")
            self.open_folder(os.path.dirname(result_file))


# 사용 예제
if __name__ == "__main__":
    consolidator = ExcelConsolidator()
    consolidator.consolidate()
    input('종료하려면 아무키나 누르세요.')