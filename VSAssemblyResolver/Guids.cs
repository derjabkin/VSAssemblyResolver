// Guids.cs
// MUST match guids.h
using System;

namespace SergejDerjabkin.VSAssemblyResolver
{
    static class GuidList
    {
        public const string guidVSAssemblyResolverPkgString = "53379d2c-e289-4f71-b241-114880543064";
        public const string guidVSAssemblyResolverCmdSetString = "a4a148d6-2f0d-4ca1-85e5-7606d51aa0d7";

        public static readonly Guid guidVSAssemblyResolverCmdSet = new Guid(guidVSAssemblyResolverCmdSetString);
    };
}