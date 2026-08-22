**********************
XlsxSharp Installation
**********************

====================
Installing XlsxSharp
====================
The easiest way to add XlsxSharp to your project is to install it using the NET CLI.

.. code-block:: batch

   C:\source> dotnet add package XlsxSharp

===============
SixLabors.Fonts
===============
XlsxSharp depends on the ``SixLabors.Fonts`` library that is only available as a beta. The legacy ``package.config``
nuget dependency resolution ignores prerelease packages during installation, resulting in the following error:

.. error::
   ::

     Unable to resolve dependency 'SixLabors.Fonts'. Source(s) used: 'nuget.org', 'Microsoft Visual Studio Offline Packages'.

To solve the error, choose one of the possible solutions:

--------------------
Use PackageReference
--------------------

Migrate from ``package.config`` to ``PackageReference``. ``package.config`` has been deprecated for 5 year and
there is a full-fledged replacement that is better in every way that works on .NET Framework. If you
use ``PackageReference`` style, the XlsxSharp is installed without the error.

See `Migrate from packages.config to PackageReference <https://learn.microsoft.com/en-us/nuget/consume-packages/migrate-packages-config-to-package-reference>`_
and related pages.

-----------------------------
Install SixLabors.Fonts first
-----------------------------
Install ``SixLabors.Fonts-beta18`` package and then install XlsxSharp package. This way, you can install the XlsxSharp
even for the ``package.config`` nuget restore.

---------------------
Use prerelease switch
---------------------

Open the *Package Manager Console* and specify the ``-IncludePrerelease`` switch during installation.

.. code-block:: batch

   PM> Install-Package XlsxSharp -Version 0.97.0 -Verbose -IncludePrerelease

==========================
Compatible implementations
==========================
XlsxSharp is a .NET Standard 2.0 library that runs on any compatible implementation (.NET Core 2.0+, .NET Framework 4.6.2, Blazor).
XlsxSharp doesn't work on Unity due to Unity `script engine <https://github.com/ClosedXML/ClosedXML/issues/1880>`_.
