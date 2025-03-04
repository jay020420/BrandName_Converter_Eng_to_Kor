import pandas as pd
from pathlib import Path
import re
import os
import glob
from collections import Counter, defaultdict

def sort_brands_by_length(brand_dict):
    """브랜드명을 길이순으로 정렬 (긴 것이 먼저 오도록)"""
    return dict(sorted(brand_dict.items(), key=lambda x: len(x[0]), reverse=True))

def load_brand_mapping(csv_file):
    """브랜드 매핑 정보를 CSV 파일에서 읽어옴"""
    try:
        current_dir = os.getcwd()
        full_path = os.path.join(current_dir, csv_file)
        print(f"현재 디렉토리: {current_dir}")
        print(f"CSV 파일 경로: {full_path}")
        
        if not os.path.exists(full_path):
            print(f"오류: {full_path} 파일을 찾을 수 없습니다.")
            return {}
            
        encodings = ['utf-8', 'cp949', 'euc-kr']
        for encoding in encodings:
            try:
                print(f"{encoding} 인코딩으로 파일 읽기 시도...")
                brands_df = pd.read_csv(full_path, encoding=encoding)
                
                if 'eng_name' not in brands_df.columns or 'kor_name' not in brands_df.columns:
                    print(f"오류: CSV 파일에 필요한 컬럼이 없습니다. 현재 컬럼: {brands_df.columns.tolist()}")
                    return {}
                
                brands_df['eng_name'] = brands_df['eng_name'].str.strip()
                brands_df['kor_name'] = brands_df['kor_name'].str.strip()
                
                duplicates = brands_df['eng_name'].duplicated()
                if duplicates.any():
                    print("경고: 다음 브랜드명이 중복되어 있습니다:")
                    print(brands_df[duplicates]['eng_name'].values)
                
                print("\nCSV 파일 읽기 성공! 데이터 미리보기:")
                print(brands_df.head())
                print(f"총 {len(brands_df)}개의 브랜드 정보를 읽었습니다.")
                
                return sort_brands_by_length(dict(zip(brands_df['eng_name'], brands_df['kor_name'])))
                
            except UnicodeDecodeError:
                print(f"{encoding} 인코딩으로 읽기 실패")
                continue
            except Exception as e:
                print(f"{encoding} 인코딩으로 읽는 중 오류 발생: {e}")
                continue
        
        print("모든 인코딩 시도 실패")
        return {}
    
    except Exception as e:
        print(f"CSV 파일 처리 중 예외 발생: {e}")
        return {}

def load_clean_text_exceptions(csv_file):
    """clean_text 예외 단어 목록을 CSV 파일에서 읽어옴"""
    try:
        current_dir = os.getcwd()
        full_path = os.path.join(current_dir, csv_file)
        print(f"\nClean text 예외 파일 경로: {full_path}")
        
        if not os.path.exists(full_path):
            print(f"오류: {full_path} 파일을 찾을 수 없습니다.")
            return set()
            
        encodings = ['utf-8', 'cp949', 'euc-kr']
        for encoding in encodings:
            try:
                print(f"{encoding} 인코딩으로 파일 읽기 시도...")
                exceptions_df = pd.read_csv(full_path, encoding=encoding)
                
                if 'word' not in exceptions_df.columns:
                    print(f"오류: CSV 파일에 'word' 컬럼이 없습니다. 현재 컬럼: {exceptions_df.columns.tolist()}")
                    return set()
                
                exceptions = set(word.strip().lower() for word in exceptions_df['word'] if pd.notna(word))
                print(f"총 {len(exceptions)}개의 예외 단어를 읽었습니다.")
                return exceptions
                
            except UnicodeDecodeError:
                print(f"{encoding} 인코딩으로 읽기 실패")
                continue
            except Exception as e:
                print(f"{encoding} 인코딩으로 읽는 중 오류 발생: {e}")
                continue
        
        print("모든 인코딩 시도 실패")
        return set()
    
    except Exception as e:
        print(f"CSV 파일 처리 중 예외 발생: {e}")
        return set()

def clean_text(text, exceptions=None):
    """텍스트 정리: 연속된 공백 제거 및 중복 단어 제거"""
    if exceptions is None:
        exceptions = set()
    
    text = re.sub(r'\s+', ' ', text.strip())
    
    # 예외 패턴 보호를 위한 임시 처리
    protected_text = text
    exception_placeholders = {}
    
    # 여러 단어로 된 예외 패턴을 임시 플레이스홀더로 대체
    for idx, exception in enumerate(sorted(exceptions, key=len, reverse=True)):
        placeholder = f"__EXCEPTION_{idx}__"
        pattern = re.compile(re.escape(exception), re.IGNORECASE)
        protected_text = pattern.sub(placeholder, protected_text)
        exception_placeholders[placeholder] = exception
    
    # 일반적인 중복 단어 제거 처리
    words = protected_text.split()
    seen_words = set()
    cleaned_words = []
    
    for word in words:
        # 플레이스홀더는 그대로 보존
        if word.startswith("__EXCEPTION_"):
            cleaned_words.append(word)
        else:
            word_lower = word.lower()
            if word_lower not in seen_words:
                cleaned_words.append(word)
                seen_words.add(word_lower)
    
    result = ' '.join(cleaned_words)
    
    # 플레이스홀더를 원래 예외 패턴으로 복원
    for placeholder, original in exception_placeholders.items():
        result = result.replace(placeholder, original)
    
    return result

def extract_english_brand(text):
    """영문 브랜드명 추출 (첫 부분의 연속된 영문 단어들)"""
    # 정리된 텍스트에서 단어들을 추출
    words = text.split()
    if not words:
        return []
    
    potential_brands = []
    current_brand = []
    
    # 첫 부분부터 연속된 영문 단어들을 찾음
    for word in words:
        # 순수하게 영문으로만 구성된 단어인지 확인 (숫자, 특수문자 제외)
        if re.match(r'^[A-Za-z]+$', word):
            current_brand.append(word)
        else:
            break
    
    # 발견된 영문 단어들로 가능한 브랜드명 조합 생성
    for i in range(1, len(current_brand) + 1):
        brand = ' '.join(current_brand[:i])
        potential_brands.append(brand)
    
    return potential_brands

def convert_product_name(original_name, brand_mapping, clean_exceptions=None):
    """개별 상품명 변환 - 한글 브랜드명 추가"""
    print(f"\n처리 중인 상품명: {original_name}")
    
    # 입력 문자열 표준화
    original_name = str(original_name).strip()
    cleaned_name = clean_text(original_name, clean_exceptions)
    print(f"정리된 상품명: {cleaned_name}")
    
    # 이미 한글 브랜드명으로 시작하는지 확인
    for kor_name in brand_mapping.values():
        if cleaned_name.lower().startswith(f"{kor_name.lower()} "):
            print(f"이미 변환된 상품명 감지: {kor_name}")
            return cleaned_name
    
    # 상품명에서 영문 브랜드 부분 추출
    potential_brands = extract_english_brand(cleaned_name)
    print(f"추출된 브랜드 후보: {potential_brands}")
    
    # 브랜드 매칭 시도 (대소문자 구분 없이)
    for potential_brand in potential_brands:
        potential_brand_lower = potential_brand.lower()
        
        for eng_name, kor_name in brand_mapping.items():
            eng_name_lower = eng_name.lower()
            
            # 정확한 매칭 또는 공백을 무시한 매칭 시도
            if (potential_brand_lower == eng_name_lower or 
                potential_brand_lower.replace(" ", "") == eng_name_lower.replace(" ", "")):
                print(f"브랜드 매칭 성공: {eng_name} -> {kor_name}")
                converted_name = f"{kor_name} {original_name}"
                return clean_text(converted_name, clean_exceptions)
    
    print("매칭되는 브랜드를 찾지 못했습니다.")
    return cleaned_name

def get_brand_part(product_name):
    """상품명에서 브랜드 부분을 추출 (영문 브랜드 추출)"""
    return extract_english_brand(product_name)

def convert_product_names(excel_path, brand_csv, clean_exceptions=None):
    """엑셀 파일의 상품명 변환"""
    try:
        current_dir = os.getcwd()
        full_excel_path = os.path.join(current_dir, excel_path)
        print(f"\n엑셀 파일 경로: {full_excel_path}")
        
        if not os.path.exists(full_excel_path):
            print(f"오류: {full_excel_path} 파일을 찾을 수 없습니다.")
            return None, [], []
        
        brand_mapping = load_brand_mapping(brand_csv)
        if not brand_mapping:
            print("브랜드 정보를 읽을 수 없습니다. CSV 파일을 확인해주세요.")
            return None, [], []
        
        print("\n엑셀 파일 읽기 시도...")
        df = pd.read_excel(full_excel_path)
        print(f"엑셀 파일 읽기 성공! 총 {len(df)}행의 데이터를 읽었습니다.")
        
        # 처음 5개 행 삭제
        df = df.iloc[5:].reset_index(drop=True)
        
        changes = []
        unmatched_records = []
        
        # 모든 한글 브랜드명을 소문자로 변환하여 set으로 저장
        korean_brands = {brand.lower() for brand in brand_mapping.values()}
        
        # D열(index 3)의 모든 행 처리
        for idx in range(len(df)):
            original_name = str(df.iloc[idx, 3])  # D열(index 3)의 값
            cleaned_name = clean_text(original_name, clean_exceptions)
            potential_brands = extract_english_brand(cleaned_name)
            converted_name = convert_product_name(original_name, brand_mapping, clean_exceptions)
            
            # 브랜드 매칭 결과 판단
            is_matched = False
            
            # 1. 변환된 경우 체크
            if original_name != converted_name:
                changes.append((original_name, converted_name))
                is_matched = True
            
            # 2. 이미 한글 브랜드로 시작하는 경우 체크
            if not is_matched and potential_brands:
                first_word = potential_brands[0].lower()
                for kor_brand in korean_brands:
                    if first_word.startswith(kor_brand):
                        is_matched = True
                        break
            
            # 매칭되지 않은 경우에만 unmatched_records에 추가
            if not is_matched and potential_brands:
                unmatched_records.append({
                    'original_name': original_name,
                    'potential_brand': potential_brands[0]
                })
            
            df.iloc[idx, 3] = converted_name
        
        output_path = Path(full_excel_path)
        new_path = output_path.parent / f"converted_{output_path.name}"
        
        # 헤더 없이 저장
        df.to_excel(new_path, index=False, header=False)
        
        print(f"\n=== {os.path.basename(excel_path)} 처리 결과 ===")
        print(f"총 {len(df)}개의 상품명 처리")
        print(f"- 변환된 상품: {len(changes)}개")
        print(f"- 이미 한글 브랜드: {len(df) - len(changes) - len(unmatched_records)}개")
        print(f"- 매칭되지 않은 상품: {len(unmatched_records)}개")
        
        return str(new_path), changes, unmatched_records
    
    except Exception as e:
        print(f"엑셀 파일 처리 중 오류 발생: {e}")
        return None, [], []

def process_multiple_files(brand_file, exceptions_file):
    """여러 개의 products 파일을 순차적으로 처리"""
    try:
        # 브랜드 매핑과 clean_text 예외 목록 로드
        brand_mapping = load_brand_mapping(brand_file)
        clean_exceptions = load_clean_text_exceptions(exceptions_file)
        
        if not brand_mapping:
            print("브랜드 정보를 읽을 수 없습니다.")
            return
            
        excel_files = glob.glob("products*.xlsx")
        
        if not excel_files:
            print("처리할 products*.xlsx 파일을 찾을 수 없습니다.")
            return
        
        print(f"발견된 파일들: {excel_files}")
        
        # 파일별 통계를 저장할 딕셔너리
        file_stats = defaultdict(dict)
        all_unmatched_records = []  # 모든 파일의 미매칭 레코드를 누적할 리스트
        
        # 각 파일 처리
        for file in sorted(excel_files):
            print(f"\n=== {file} 처리 시작 ===")
            converted_file, changes, unmatched_records = convert_product_names(
                file, brand_file, clean_exceptions=clean_exceptions
            )
            if converted_file:
                print(f"파일 변환 완료: {converted_file}")
                
                # 파일별 통계 저장
                file_stats[file] = {
                    'total_changes': len(changes),
                    'total_unmatched': len(unmatched_records),
                    'unmatched_records': unmatched_records
                }
                
                # 전체 미매칭 레코드 누적
                all_unmatched_records.extend(unmatched_records)
            else:
                print(f"{file} 처리 실패")
            
            print(f"=== {file} 처리 완료 ===\n")
        
        # 미매칭 브랜드 결과를 저장할 리스트
        unmatched_brands_results = []
        
        # 모든 미매칭 레코드에 대해 브랜드별로 집계
        brand_counter = defaultdict(list)
        for record in all_unmatched_records:
            if re.match(r'^[A-Za-z\s]+$', record['potential_brand']):
                brand = record['potential_brand'].lower().strip()
                brand_counter[brand].append(record['original_name'])
        
        # 결과를 DataFrame으로 변환
        for brand, original_names in brand_counter.items():
            unmatched_brands_results.append({
                '브랜드명': brand,
                '발생횟수': len(original_names),
                '전체상품명': ' | '.join(original_names)  # 원본 상품명들을 | 로 구분하여 저장
            })
        
        # DataFrame 생성 및 정렬
        results_df = pd.DataFrame(unmatched_brands_results)
        if not results_df.empty:
            results_df = results_df.sort_values('발생횟수', ascending=False)
            
            # CSV 파일로 저장
            output_filename = 'unmatched_brands_detailed.csv'
            results_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            print(f"\n미매칭 브랜드 정보가 '{output_filename}' 파일로 저장되었습니다.")
            print(f"총 {len(results_df)}개의 미매칭 브랜드가 기록되었습니다.")
            
            # 상위 20개 브랜드 출력
            print("\n=== 전체 미매칭 브랜드 통계 (상위 20개) ===")
            top_20 = results_df.head(20)
            for _, row in top_20.iterrows():
                print(f"- {row['브랜드명']}: {row['발생횟수']}회")
        else:
            print("\n미매칭 브랜드가 없습니다.")
            
    except Exception as e:
        print(f"파일 처리 중 오류 발생: {e}")

if __name__ == "__main__":
    brand_file = "brands.csv"
    exceptions_file = "exceptions.csv"  # clean_text 예외 단어 목록 파일
    process_multiple_files(brand_file, exceptions_file)