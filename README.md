특정 엑셀 파일의 찾으려는 값이 실제 디스크에 존재하는지 검사하는 스크립트입니다.

1. 경로 선택 버튼 클릭해 경로 지정 (self.path에 저장 / file = self.path + path_table)
2. 찾으려는 테이블 버튼 클릭

아래 변수에 상황에 맞게 값을 입력해 사용합니다.

path_table = ""  # 엑셀 데이터파일 경로
path_target = ""  # 파일 존재를 확인할 경로
self.check_table(path_table, path_target, "N", 5)   # N열 5행부터 N열 탐색, 찾으려는 범위에 맞게 변경

실행 예시

![image](https://github.com/Entournure/FileCheckXLS/assets/50042686/4add705f-eb18-4ac3-8877-4137a6d5b311)
