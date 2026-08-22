**********************
XlsxSharp Installation
**********************

====================
Installing XlsxSharp
====================
The easiest way to add XlsxSharp to your project is to install it using the .NET CLI.
XlsxSharp is currently published as a prerelease, so the ``--prerelease`` switch is required
until a stable version is released.

.. code-block:: batch

   C:\source> dotnet add package XlsxSharp --prerelease

You can also install the package from the *Package Manager Console* in Visual Studio.

.. code-block:: batch

   PM> Install-Package XlsxSharp -IncludePrerelease

==========================
Compatible implementations
==========================
XlsxSharp targets ``net10.0`` and runs on .NET 10 or later, on every platform supported by
.NET (Windows, Linux, macOS).

.. note::
   XlsxSharp is a fork of `ClosedXML <https://github.com/ClosedXML/ClosedXML>`_ whose primary goal
   is to keep the library current with the latest .NET releases. Unlike ClosedXML, it does not
   target .NET Standard 2.0 or .NET Framework. If you need to run on those platforms, use
   ClosedXML instead.

XlsxSharp doesn't work on Unity due to Unity `script engine <https://github.com/ClosedXML/ClosedXML/issues/1880>`_.
