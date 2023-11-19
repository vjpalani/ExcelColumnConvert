This project provides a versatile C# function designed to seamlessly convert Excel Sheet Column Letters to Column Names and vice versa. Implementing the hexagonal architecture, this solution ensures flexibility for integration into various contexts, including API or console applications.

Installation
To seamlessly integrate this function into your project, proceed with the following steps:

Clone this Repository:

sh
Copy code
git clone https://github.com/vjpalani/ExcelColumnConvert.git
Build the Solution:

sh
Copy code
cd excel-column-convert
dotnet build
Run API Server:

sh
Copy code
cd Api
dotnet run
Now, you can make API calls using tools like curl or your preferred API testing tool.

Example API Call:

sh
Copy code
curl --location 'http://localhost:5100/ExcelConverter/columnConverter' \
--header 'Content-Type: application/json' \
--data '{
    "columnLetter":"AA"
}'
Run API Unit Tests:

sh
Copy code
cd ../Tests/Api.Tests
dotnet test
Run Console Application:

sh
Copy code
cd ..
cd console
dotnet run
Follow the on-screen instructions to enter either a Column Name or Column Letter for conversion.

Run Console Application Tests:

sh
Copy code
cd ../Tests/console.Tests
dotnet test
Run Core Function Tests:

sh
Copy code
cd ..
cd core.Tests
dotnet test
Hexagonal Architecture Explanation
The hexagonal architecture distinguishes core business logic from external dependencies. The ExcelColumnConverter class encapsulates the core logic for Excel column conversion, ensuring modularity. External dependencies, such as input/output or API interactions, can be seamlessly integrated without impacting the core logic.

Feel free to explore and enhance the project as needed for your specific use cases. Unit tests are provided for the API, console, and core functions, ensuring the robustness of the solution across various scenarios.
