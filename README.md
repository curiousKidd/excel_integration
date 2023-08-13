## 🐳페이코 엑셀 파일 추출기
### 문제 상황

기존에는 회사 내부에서 페이코 통합 사용 리스트를 전달받아, 이를 수작업으로 페이코 승인 내역과 매칭시키고 취합하는 과정에서 많은 시간이 소모되었습니다.

이런 상황에서 효율적으로 업무를 처리할 방법이 필요했습니다.

### 프로젝트 목표

1. 페이코 관리자 페이지에서 다운로드할 수 있는 페이코 사용 내역 리스트를 읽어와서 가공합니다.

2. 페이코 승인 내역과 통합 사용자들을 자동으로 매칭하여 사용자별 식대 사용 금액을 취합합니다.

3. 자동으로 생성된 통합 사용자 리스트를 제공합니다.

### 프로젝트 이점

1. 시간 절약 : 프로젝트를 통해 수작업으로 하던 업무를 자동화하므로, 업무 처리 시간을 현저히 줄일 수 있습니다.

2. 정확성 향상 : 개인들의 수기 작성으로 실수가 발생할 가능성이 있으나, 프로젝트의 사용으로 정확성을 향상 시킬 수 있습니다.

3. 편의성 : 통합 사용자 리스트를 자동으로 생성하므로, 팀 구성원들은 식대 사용 내역을 신속하게 확인할 수 있습니다.

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

### 2023.08.07
``` java 
private String getNames(PaycoDTO paycoDTO, int price, List<PaycoDTO> paycoDTOS) {
    HashMap<String, Integer> map = new HashMap<>();
    StringBuilder sb = new StringBuilder();

    paycoDTOS.stream()
            .filter(f ->
                    f.getName().equals(paycoDTO.getName())
                            && f.getTranDate().substring(0, 10).equals(paycoDTO.getTranDate().substring(0, 10))
                            && f.getTicketType().equals(paycoDTO.getTicketType())
                            && !f.getTranType().equals("취소")
            )
						/* 20230807 이전 로직 */
            //.map(fm -> !fm.getTranType().equals("승인")
            //        ? fm.getTicketPaymentAmount() > price
            //       ? sb.append(overPriceName(fm, paycoDTOS))
            //        : fm.getTranType().equals("받기")
            //        ? map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) + fm.getTicketPaymentAmount())
            //        : map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) - fm.getTicketPaymentAmount())
            //        : paycoDTO.getTranNumber().equals(fm.getTranNumber())
            //        ? map.put(fm.getName(), fm.getTicketPaymentAmount()) : "")
						.map(fm -> {
                    if (!"승인".equals(fm.getTranType())) {
                        if (fm.getTicketPaymentAmount() > price) {
                            sb.append(overPriceName(fm, paycoDTOS));
                        } else if ("받기".equals(fm.getTranType())) {
                            map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) + fm.getTicketPaymentAmount());
                        } else {
                            map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) - fm.getTicketPaymentAmount());
                        }
                    } else if (paycoDTO.getTranNumber().equals(fm.getTranNumber())) {
                        map.put(fm.getName(), fm.getTicketPaymentAmount());
                    }
                    return ""; // 빈 문자열 반환 (변경 가능성 있음)

            .collect(Collectors.toList());

    for (String key : map.keySet()) {
        if (map.get(key) > 0) {
            sb.append(key);
            sb.append(", ");
        }
    }

    return sb.toString();
}
```
20230807 기준으로 .map() 안에 로직이 변경되었습니다  
→ 기존 삼항식을 사용하던 로직에서 if문을 사용하도록 로직 수정
- 삼항식을 다중으로 사용하다보니 오히려 가독성이 떨어지게 되어 if문으로 수정되었습니다.

