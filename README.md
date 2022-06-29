# Excel DNA Addin test

Working Excel-DNA + .NET 6 + Google OR-Tools Excel Add-in example.

Compared to a "vanilla" project, I changed the target from net6.0 to net6.0-windows and added a post-script build to copy the C++ native DLL to the build folder. Without this, Excel can't find the DLL.
