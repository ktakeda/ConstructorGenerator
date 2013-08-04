// Guids.cs
// MUST match guids.h
using System;

namespace ktakeda.ConstructorGenerator
{
    static class GuidList
    {
        public const string guidConstructorGeneratorPkgString = "70b3dc53-a5d3-4cb2-b9f1-6a89b00ae46f";
        public const string guidConstructorGeneratorCmdSetString = "4dd90096-2911-464c-8ab8-7aa60c21df90";

        public static readonly Guid guidConstructorGeneratorCmdSet = new Guid(guidConstructorGeneratorCmdSetString);
    };
}