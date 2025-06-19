@echo off
REM Test script for Windows PowerShell/Command Prompt
REM This script shows how to properly test the token authentication

echo 🧪 Lokka MCP Server - Token Test Instructions
echo ================================================

echo.
echo 📋 Method 1: Set environment variable then run test
echo.
echo In PowerShell:
echo   $env:ACCESS_TOKEN = "your-jwt-token-here"
echo   npm run test:simple
echo.
echo In Command Prompt:
echo   set ACCESS_TOKEN=your-jwt-token-here
echo   npm run test:simple
echo.

echo 📋 Method 2: One-line PowerShell command
echo.
echo   $env:ACCESS_TOKEN = "your-jwt-token"; npm run test:simple
echo.

echo 📋 Method 3: Use the interactive demo to get a fresh token
echo.
echo   npm run demo:token
echo.

echo ⚠️  Important Notes:
echo - Your token appears to be valid (it's a proper JWT format)
echo - The token may have expired (check the 'exp' claim)
echo - Ensure the token has Microsoft Graph permissions
echo - Don't include the token directly in the npm command line
echo.

echo 💡 Quick test with your token:
echo Copy and paste this command in PowerShell:
echo.
echo $env:ACCESS_TOKEN = "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImJ0YXU0SV83LWlZTGVweFlGX1dkZWtpOEFUYW5xOUp1QzRO[Truncated for example]"; npm run test:simple
echo.
echo This will test the token authentication functionality!
pause
