# Excel Export Tool - 기획 데이터 자동 변환 및 배포 도구

**ExcelExport**는 게임 기획 데이터(Excel 파일)를 자동으로 추출, 변환,  
output 파일을 생성하고 이를 게임 서버에 배포하는 자동화 도구입니다.  
기획자의 수동 작업을 최소화하고, Jenkins 및 Slack과 연동하여  
라이브 서비스 배포 시 안정성과 효율성을 높이도록 설계되었습니다.

---

## 주요 기술 구성

- C# (.NET Framework or .NET Core)
- ClosedXML / ExcelDataReader (Excel 처리)
- JSON/CSV Export

---

## 주요 기능

### Excel 데이터 자동 변환
- 다수의 Excel 시트를 파싱
- 지정된 포맷으로 JSON/CSV 파일 출력
- 오류/포맷 예외 자동 검출 및 로그 출력

### Jenkins 기반 자동 실행
- 배포 서버에서 Jenkins 빌드 트리거 시 자동 실행
- 별도 빌드/실행 없이 CLI에서 직접 수행 가능

### 기획자용 개별 실행
- 기획자가 직접 실행할 수 있도록 단일 실행 파일 제공
- 파일 경로 지정만으로 사용 가능 (명령행 인자 또는 환경설정)
