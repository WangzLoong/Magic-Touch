@echo off

:START
set /p OPENAI_API_KEY=Please enter your OpenAI API key:

if not "%OPENAI_API_KEY:~0,3%"=="sk-" (
    echo Error: OpenAI API key must start with "sk-".
    goto START
)

setx OPENAI_API_KEY %OPENAI_API_KEY% /m

IF "%OPENAI_API_KEY%"=="" (
   ECHO OpenAI API Key not found in environment variables) ELSE (
   ECHO OpenAI API Key successfully added to environment variables.
)


pause