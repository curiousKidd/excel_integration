## 페이코 엑셀 파일 추출기
```
현 회사에서 식대를 `페이코`를 통하여 제공받는 중입니다.  
페이코 어플을 통해서 `식권`으로 포인트를 제공받고 있으며, 제휴맺은 식당에서 결제가 가능합니다.  
어플 내부에서 `페이코 포인트`를 `타인에게 이전`하여 `사용`할 수 있습니다.

현 회사 지인 업무중 하나가 페이코 사용 내역을 취합하는 것이었습니다.  
매달 팀 내부에서 페이코를 이전하여 사용한 내역의 한해서 리스트를 취합하였고,   
전달 받은 파일을 기준으로 사용내역과 통합 사용 내용을 취합하는 프로젝트입니다.

java stream을 공부중이었기에 java8을 사용하였습니다.
```

## ✅payco project
- MVC
- java8
- org.apache.poi -> 엑셀파일 read lib
```
기본 로직을 테스트 하기 위한 프로젝트 입니다.
단순 MVC 구조를 사용하였으며, 입력한 엑셀 파일과 출력해야 할 엑셀 데이터를 화면에 출력하는 용도로 사용하였습니다.
```

## ✅payco_swing project
- java8
- java swing
- org.apache.poi

payco project를 통해서 출력할 엑셀 데이터를 확인하였습니다.  
해당 로직을  exe 파일을 만들기 위해서 `java swing`을 사용하여 gui 프로젝트를 구성하였습니다.   
java swing을 사용한 이유는 주변 사람들의 추천 및 보다 쉽게 정보를 수집 할 수 있다는 장점이 있어 사용하였습니다.  

해당 프로젝트에서는 페이코 관리자페이지에서 다운받은 사용 내역 엑셀파일`(payco Excel)`과   
회사 내부에서 작성된 통합 사용 엑셀 파일들(부서별 작성, `bay Excel`)을 기준으로 엑셀 데이터를 생성하였습니다.

<img src="https://github.com/curiousKidd/excel_integration/assets/77770531/7cd03326-7335-4aa0-aa15-137813e7a9fd" width="500" height="250"/>

``` text
bay Excel 의 경우 다중으로 선택하여 파일들을 한번에 입력 가능하도록 설정하였습니다. 
->  fileComponent.setMultiSelectionEnabled(true); // 다중 선택 
```  
<img src="https://github.com/curiousKidd/excel_integration/assets/77770531/2f4c5e9f-a51f-4d72-be93-1997160a0957" width="500" height="240"/>


## ✅payco_swing_new project
- java8
- java swing
- org.apache.poi


payco_swing 프로젝트에서의 단점을 보완 및 개선한 프로젝트입니다.  
기존 `bay Excel`과 `payco Excel`을 모두 입력 받아야했던 점을 보완하여 `payco Excel`만 입력받도록 수정하였습니다.

<img src="https://github.com/curiousKidd/excel_integration/assets/77770531/8987def7-b719-4452-a11e-caea9524a85a" width="500" height="250"/>

### 2023.07
회사 내부 페이코 금액이 증가되었습니다. 이전 금액을 기준으로 통합사용자를 추출하였으나,
금액이 증가된 이후 수정이 필요하였습니다. 이후 금액을 직접 작성할 수 있도록 gui에 text 입력칸을 추가하였습니다.

<img src="https://github.com/curiousKidd/excel_integration/assets/77770531/47434510-3ac7-43e5-9560-a95692532e00" width="500" height="250"/>

### 2023.07.19
사용자(현 직장 동료)의 의견으로 수정 진행 예정
기준 금액보다 많은 금액을 사용한 사람들의 이름은 무조건 나올 수 있는지 확인 요청


