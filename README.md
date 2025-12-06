# VintaSoft Console eSCL Image Scanning Demo

This .NET console project uses <a href="https://www.vintasoft.com/vstwain-dotnet-index.html">VintaSoft TWAIN .NET SDK</a> and demonstrates how to acquire images from eSCL image scanner in console application in Windows and Linux.


## Usage
1. Get the 30 day free evaluation license for <a href="https://www.vintasoft.com/vstwain-dotnet-index.html" target="_blank">VintaSoft TWAIN .NET SDK</a> as described here: <a href="https://www.vintasoft.com/docs/vstwain-dotnet/Licensing-Twain-Evaluation.html" target="_blank">https://www.vintasoft.com/docs/vstwain-dotnet/Licensing-Twain-Evaluation.html</a>

2. Update the evaluation license in "CSharp\Program.cs" file:
   ```
   Vintasoft.EsclImageScanning.EsclGlobalSettings.Register("REG_USER", "REG_EMAIL", "EXPIRATION_DATE", "REG_CODE");
   ```

3. Build the project ("EsclImageScanningConsoleDemo.Net10.csproj" file) in Visual Studio or using .NET CLI:
   ```
   dotnet build EsclImageScanningConsoleDemo.Net10.csproj
   ```

4. Run compiled application.


## Documentation
VintaSoft TWAIN .NET SDK on-line User Guide and API Reference for .NET developer is available here: https://www.vintasoft.com/docs/vstwain-dotnet/


## Support
Please visit our <a href="https://myaccount.vintasoft.com/">online support center</a> if you have any question or problem.
