#pragma once

using namespace System;
using namespace System::Collections::Generic;

namespace sharedData
{
    public ref class ListStorage
    {
    public:
        static List<List<String^>^>^ imagePaths = gcnew List<List<String^>^>();
    public:
        static List<System::String^>^ appTypes = gcnew List<System::String^>();
    };
}