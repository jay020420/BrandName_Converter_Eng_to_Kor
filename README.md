스마트스토어 상품명 치환 코드입니다.

스마트스토어가 아닌 엑셀을 통하여 대량 수정이 가능한 마켓에서 모두 사용가능합니다.

주로 대량 등록으로 브랜드 이름 번역 없이 업로드 되는 셀러들을 위해 작성하였습니다.

코드 실행 시 코드 파일 (.py)이 위치한 폴더와 같은 폴더 내에 있는 엑셀(.xlsx) 중 products로 시작하는 모든 파일을 인식합니다.

이후 brand.csv에 기입된 정보에 따라 번역됩니다.
exception.csv에 기입된 브랜드들은 따로 번역하지 않습니다.

예시로
products(1).xlsx와 products(2).xlsx가 존재하는 경우
products(1).xlsx 처리 이후 products(2).xlsx 등 순차적으로 처리합니다.

기본적으로 다음과 같이 번역됩니다.
[영문명] 상품명
[한글명] [영문명] ~상품명~

이는 코드 수정을 통해 얼마든지 변경 가능한 사항이니 자유롭게 수정하셔도 됩니다.

brands.csv는 주로 명품과 컨텀포러리 브랜드 명으로 이루어져 있습니다.
차후 서버를 구축하여 미번역된 브랜드들을 유저가 업로드할 수 있도록 하겠습니다.

현재 여러 개의 products.xlsx 파일을 변환하는 경우 오류 집계가 되지 않고 있습니다.
사용에 참고해주시길 바랍니다.
