#!/bin/bash

# Demo version
MSBuild.exe ../WordLinkValidator.sln /property:Configuration=Release -p:DefineConstants=IS_DEMO

# Full version
MSBuild.exe ../WordLinkValidator.sln /property:Configuration=Release
