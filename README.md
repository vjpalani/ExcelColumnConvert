Excel Column Converter
======================

This project presents a versatile C# function crafted to effortlessly convert Excel Sheet Column Letters to Column Names and vice versa. By employing the hexagonal architecture, this solution guarantees flexibility for integration into diverse contexts, including API or console applications.

Installation
------------

### 1\. Clone this Repository:

```
git clone https://github.com/vjpalani/ExcelColumnConvert.git
```

### 2\. Build the Solution:

```
dotnet build
```

### 3\. Run API Server:

```
cd Api
dotnet run
```

Now, you can make API calls using tools like curl or your preferred API testing tool.

### 4\. Example API Call:

```
curl --location 'http://localhost:5100/ExcelConverter/columnConverter'
--header 'Content-Type: application/json'
--data '{ "columnLetter":"AA" }'
```

### 5\. Run API Unit Tests:

```
cd ../Tests/Api.Tests
dotnet test
```

### 6\. Run Console Application:

```
cd ../console
dotnet run
```

Follow the on-screen instructions to enter either a Column Name or Column Letter for conversion.

### 7\. Run Console Application Tests:

```
cd ../Tests/console.Tests
dotnet test
```

### 8\. Run Core Function Tests:

```
cd ../core.Tests
dotnet test
```

Hexagonal Architecture Explanation
----------------------------------

The hexagonal architecture differentiates core business logic from external dependencies. The `ExcelColumnConverter` class encapsulates the core logic for Excel column conversion, ensuring modularity. External dependencies, such as input/output or API interactions, can be seamlessly integrated without impacting the core logic.
