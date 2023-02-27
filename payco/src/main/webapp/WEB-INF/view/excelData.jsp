<!DOCTYPE html>
<html lang="en" xmlns:th="http://www.thymeleaf.org">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>
<div style="border: 1px solid #222;">
    <table class="tb_col">
        <tr th:each="excelData : ${excel}">
            <td th:text="">${excelData}</td>
            <%--            <td th:text="${product.name}"></td>--%>
            <%--            <td th:text="${product.price}"></td>--%>
            <%--            <td th:text="${product.quantity}"></td>--%>
        </tr>
    </table>
</div>
</body>
</html>