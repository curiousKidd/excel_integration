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
    <form th:action="@{/excel}" method="post" enctype="multipart/form-data">
        <ul>
            <li>첨부파일<input th:type="file" name="fileUpload"></li>
        </ul>
        <input th:type="submit" th:value="전송">
    </form>
</div>
</body>
</html>