- [Document Builder samples](#document-builder-samples)
  - [About](#about)
  - [Project generator: configure.py](#project-generator-configurepy)
  - [Running C++ samples](#running-c-samples)
    - [Visual Studio](#visual-studio)
    - [Qt](#qt)
    - [Makefile](#makefile)
  - [Running C# samples](#running-c-samples-1)
    - [Visual Studio](#visual-studio-1)
  - [Running Python samples](#running-python-samples)
  - [Running Java samples](#running-java-samples)

# Document Builder samples
## About
Here you can find some code samples for Document Builder library in different programming languages:
 1. [C++](#running-c-samples)
 2. [C# (.NET)](#running-c-samples-1)
 3. [Python](#running-python-samples)
 4. [Java](#running-java-samples)

## Project generator: configure.py
For running C++ and C# code samples use python script `configure/configure.py` which is able to generate:
 + Visual Studio project files
 + Qt project file
 + Makefile

To use `configure.py` you need to specify following options:
 1. Which project files to generate: `--vs`, `--qt` or `--make`. Several options are available at the same time, but some of them are not supported on all platforms. In case you provide none of these options, all available projects will be generated.
 2. Test samples with `--test TEST`. Some available options:
    - `--test all` – generate projects for both C++ and C#.
    - `--test cpp` – generate projects only for C++ samples
    - `--test cs` – generate projects only for C# samples.
    - `--test cpp/creating_basic_form` – generate only project for the specified sample.

    Several test options are available at the same time. To see all available `TEST` options call `configure.py -l`.
 3. Directory to the Document Builder with `--dir DIR`. If Document Builder is not installed in default path you have to provide path to it.

Generated files will be located in the `out` directory inside of the corresponding test folders.

## Running C++ samples

### Visual Studio

> **NOTE:** Only available on Windows

 1. Use `configure.py` to generate VS project files. For example:

    ```shell
    python configure.py --vs --test cpp/creating_basic_form --test cpp/creating_advanced_form
    ```

 2. Open `.sln` file in Visual Studio. It will prompt you to retarget Windows SDK and VS toolset to your installed version – click "OK".
 3. The solution is ready to be built and run. Documents will be created in the project files directory.

### Qt
 1. Use `configure.py` to generate Qt project files. For example:

    ```shell
    python configure.py --qt --test cpp
    ```

 2. Open `.pro` file in Qt Creator.
 3. The project is ready to be built and run. Documents will be created in the `build` directory.

### Makefile

> **NOTE:** Only available on Linux and Mac OS.

 1. Use `configure.py` to generate Makefile. For example:

    ```shell
    python configure.py --make --test cpp/filling_spreadsheet
    ```

 2. Go to the directory with generated Makefile:

    ```shell
    cd ../out/cpp/filling_spreadsheet
    ```

 3. Call

    ```shell
    make
    ```

    `make` will build and run the executable. Documents will be created in the same directory as Makefile is.

## Running C# samples

> **NOTE:** Only available on Windows with Visual Studio and .NET SDK installed

### Visual Studio
 1. Use `configure.py` to generate VS project files. For example:

    ```shell
    python configure.py --vs --test cs
    ```

 2. Open `.sln` file in Visual Studio. Depending on your installed .NET SDK version you may need to set different target framework by setting it in Visual Studio project properties or editing it directly in the `.csproj` file.
 3. The solution is ready to be built and run. Documents will be created in the project files directory.

## Running Python samples

 1. Execute `configure.py` with `--test python`:

    ```shell
    python configure.py --test python
    ```
    Provide directory to Document Builder with `--dir` if necessary.

    It is needed for generating auxiliary module containing the path to Document Builder.
 2. Go to test directory:

    ```shell
    cd python/creating_basic_form
    ```
 3. Run the script

    ```shell
    python main.py
    ```

    Documents will be created in the test directory.

## Running Java samples

> **NOTE:** JDK 8 or newer is required

 1. Go to test directory:

    ```shell
    cd java/creating_presentation
    ```

 2. Compile the `Program.java` providing the path to ***docbuilder.jar***, located in the Document Builder directory:

    ```shell
    javac -cp "C:\Program Files\ONLYOFFICE\DocumentBuilder\docbuilder.jar" Program.java
    ```

    Note that some Java samples use `json_simple` library for parsing JSON files. The library is located at *resources/utils/json/json_simple-1.1.jar*. To compile these samples, you also need to add the path to `json_simple` library under `-cp` option:

    ```shell
    javac -cp "C:\Program Files\ONLYOFFICE\DocumentBuilder\docbuilder.jar;..\..\resources\utils\json\json_simple-1.1.jar" Program.java
    ```

    Also note that on UNIX systems the path separator is `:` instead of `;`. Thus, on Linux or Mac OS it should be:

    ```shell
    javac -cp "/opt/onlyoffice/documentbuilder/docbuilder.jar:../../resources/utils/json/json_simple-1.1.jar" Program.java
    ```

 3. `.class` file should appear in the directory. Run the program with `java` and specify the same libraries under `-cp` option as in the compilation step, but also add the current directory `.`:

    ```shell
    java -cp "C:\Program Files\ONLYOFFICE\DocumentBuilder\docbuilder.jar;." Program
    ```

    Or, if the sample uses the JSON library:

    ```shell
    java -cp "C:\Program Files\ONLYOFFICE\DocumentBuilder\docbuilder.jar;..\..\resources\utils\json\json_simple-1.1.jar;." Program
    ```

    Again, the same applies to UNIX systems, but with `;` replaced with `:`.

 3. Documents will be created in the test directory.
