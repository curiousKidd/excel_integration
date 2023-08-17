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

### 2023.08.14
기존 로직의 변경이 있었습니다.

사용자 : A,B,C,D  
  
C > B > A(최종사용자) 의 이체순서가 있을 경우 A의 결제건에는 A,B,C가 있어야합니다.  
이후 D하테 A가 금액을 받을 경우에는 잔여 금액의 따라 A,D가 나와야할지 A,B,C,D가 출력되어야할지 결정되는 로직이 추가되었습니다.
또한 결제 여부에 따라 통합사용자명 출력의 기준이 변경되었습니다.  
  
기존 > `결제여부와 상관없이` 이체 받은 금액이 +일 경우에만 통합사용자의 포함되었습니다.  
변경 > `결제 시간`의 따라서 `이체를 받은 후 결제를 진행`했을때 이체받은 금액을 사용하지 않았더라도 이름이 포함되도록 변경되었습니다.



``` java 
private static List<PaycoDTO> paycoExcelData = new ArrayList<>();
private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy.MM.dd HH:mm:ss");
private static int price = 0;
```
기존 사용중이던 `paycoExcelData`, `price`와 `DATE_TIME_FORMATTER`를 `static` 변수로 변경하였습니다.  
`price`, `paycoExcelData` 같은 경우 실행 후 값이 입력될 것이기 때문에 `final`을 사용하지 않았습니다.


``` java
excelDTOS = paycoExcelData.stream()
                .sorted(compare)
                .filter(f -> set.add(f.getTranNumber()) && "승인".equals(f.getTranType()))
                .map(m -> ExcelDTO.builder()
                        .tranNumber(m.getTranNumber())
                        .tranDate(m.getTranDate())
                        .companyNumber(m.getCompanyNumber())
                        .name(m.getName())
                        .usePlace(m.getUsePlace())
                        .paymentType(m.getPaymentType())
                        .tranType(m.getTranType())
                        .totalPaymentAmount(m.getTotalPaymentAmount())
                        .ticketPaymentAmount(m.getTicketPaymentAmount())
                        .ticketType(m.getTicketType())
                        //                        .names(m.getTicketPaymentAmount() > price ? getNames(m, price, paycoDTOS) : "")
                        .names(getNames(m))
                        .build()
                )
                
                

```
m.getTicketPaymentAmount() > price 처럼 기존 price보다 높은 수치가 들어왔을때만 통합사용자를 가져오던 로직이 수정되었습니다.  
결제금액의 상관없이 이체를 받고, 이후 다시 돌려주지 않은 상태에서 결제가 발생할 경우 사용금액의 상관없이 통합사용자로 포함됩니다.

```java
    // 통합 사용자 이름 가져오기
    // paycoDTO : 승인 데이터
    private String getNames(PaycoDTO paycoDTO) {
        HashMap<String, Integer> map = new HashMap<>();
        StringBuilder sb = new StringBuilder();

        paycoExcelData.stream()
                .filter(f ->
                        f.getName().equals(paycoDTO.getName())
                                && f.getTranDate().substring(0, 10).equals(paycoDTO.getTranDate().substring(0, 10))
                                && f.getTicketType().equals(paycoDTO.getTicketType())
                                && !"취소".equals(f.getTranType())
                )
                // .map() 로직 수정이 발생하였습니다.
                // 받기건의 경우에는 무조건적으로 또다른 이체 이력이 존재하는지 확인하는 로직 추가
                .map(fm -> {
                    if (!"승인".equals(fm.getTranType())) {
                        if ("받기".equals(fm.getTranType())) {
                            map.putAll(amountTracking(fm));
                            //                            map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) + fm.getTicketPaymentAmount());
                        } else {
                            map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) - fm.getTicketPaymentAmount());
                        }
                    } else if (paycoDTO.getTranNumber().equals(fm.getTranNumber())) {
                        map.put(fm.getName(), map.getOrDefault(fm.getName(), 0) + fm.getTicketPaymentAmount());

                    }
                    return fm;
                })
                .collect(Collectors.toList());

        Map<String, Integer> filteredMap = map.entrySet()
                .stream()
                .filter(f -> f.getValue() > 0)
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));

        if (filteredMap.size() > 1) {
            for (String key : filteredMap.keySet()) {
                sb.append(key);
                sb.append(", ");
            }
        }

        return sb.toString();
    }
```

```java
    // 받기 금액 추적
    // PaycoDTO dto : 받기 내역 DTO
    private Map<String, Integer> amountTracking(PaycoDTO dto) {
        HashMap<String, Integer> map = new HashMap<>();
        StringBuilder sb = new StringBuilder();
        AtomicBoolean approvalCheck = new AtomicBoolean(false);

        Comparator<PaycoDTO> compare = Comparator
                .comparing(PaycoDTO::getTicketType)
                .thenComparing(PaycoDTO::getTranNumber).reversed();
        
        // 현재 받은 데이터를 무조건적으로 포함합니다. -> 어떠한 승인건의 받기 내역
        map.put(dto.getUsePlace(), dto.getTicketPaymentAmount());
        
        // 데이터를 정렬하는 조건이 추가되었습니다.
        // 현재 받기건과 같은 날짜의 데이터를 조회하고, 현 데이터보다 이전의 데이터들만 가져오기로 수정되었습니다.
        paycoExcelData.stream()
                .sorted(compare)
                .filter(f ->
                        f.getName().equals(dto.getUsePlace())
                                && f.getTranDate().substring(0, 10).equals(dto.getTranDate().substring(0, 10))
                                && LocalDateTime.parse(f.getTranDate(), DATE_TIME_FORMATTER).compareTo(LocalDateTime.parse(dto.getTranDate(), DATE_TIME_FORMATTER)) < 0
                                && f.getTicketType().equals(dto.getTicketType())
                                && !"취소".equals(f.getTranType())
                )
                // 받기건의 경우 해당 method를 재호출하여, 내가 다른사람에게 받은건 없는지, 체크합니다.
                // 승인건이면서 가격이 price로 딱 맞아나눠지면(잔액이 남지 않는 결제의 경우) 이전 받기 사용자를 포함하지 않도록 수정되었습니다.
                .map(m -> {
                    if (!approvalCheck.get()) {
                        if ("받기".equals(m.getTranType())) {
                            map.putAll(amountTracking(m));
                            //                        map.put(m.getUsePlace(), map.getOrDefault(m.getUsePlace(), 0) + m.getTicketPaymentAmount());
                        } else if ("보내기".equals(m.getTranType())) {
                            map.put(m.getUsePlace(), map.getOrDefault(m.getUsePlace(), 0) - m.getTicketPaymentAmount());
                        } else if ("승인".equals(m.getTranType()) && m.getTicketPaymentAmount() % price == 0) {
                            approvalCheck.set(true);
                        }
                    }
                    return m;
                })
                .collect(Collectors.toList());

        //Map filter
        Map<String, Integer> filteredMap = map.entrySet()
                .stream()
                .filter(f -> f.getValue() > 0)
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));

        return filteredMap;
    }
```