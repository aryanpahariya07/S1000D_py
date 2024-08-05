<!DOCTYPE html>
<html>
<head>
    <title>File Upload Form</title>
</head>
<body>
    <h2>Upload .docx File</h2>
    <form action="/" method="post" enctype="multipart/form-data">
        Select file to upload:
        <input type="file" name="file" id="file">
        <br><br>
        <input type="submit" value="Upload File" name="submit">
    </form>
</body>
</html>
