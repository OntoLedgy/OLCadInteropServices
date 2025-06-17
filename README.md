# OLCADServices

OLCADServices is a comprehensive .NET interoperability library for working with CAD applications, specifically AutoCAD and MicroStation. This library enables programmatic access to CAD documents, allowing for reading, manipulation, and exportation of CAD entities and document information.

## Features

- **Multi-CAD Support**: Works with both AutoCAD and MicroStation
- **Entity Extraction**: Read and process various CAD entities (lines, circles, text, etc.)
- **Data Export**: Export CAD data to CSV and JSON formats
- **Layer Management**: Access and report on document layers
- **Block/Reference Analysis**: Analyze and report on block references and nested components
- **Layout Management**: Access document layouts and paper spaces

## Prerequisites

The library requires:

- .NET 8.0 or higher
- AutoCAD 2021 (for AutoCAD interop)
- MicroStation 2024 (for MicroStation interop)
- STA (Single-Threaded Apartment) thread state for COM interop operations

## Architecture

### Core Components

The solution consists of two main projects:

1. **CADInteropServices**: The main library providing CAD interoperability
2. **TestCADInterOpServices**: Test project demonstrating usage and providing examples

### Project Structure

```
CADInteropServices/
├── Factories/            # Entity creation factories
├── Objects/
│   ├── AutoCAD/          # AutoCAD-specific objects
│   │   ├── Annotations/  # Text, leaders, hatches
│   │   ├── Shapes/       # Lines, circles, arcs, etc.
│   │   └── Spaces/       # Model space, paper space, layouts
│   └── MicroStation/     # MicroStation-specific objects
├── Orchestrators/        # High-level operation coordinators
└── Transformers/         # Coordinate and transformation utilities
```

### Design Patterns

The library implements several design patterns:

- **Facade Pattern**: Orchestrator classes provide simplified interfaces for complex operations
- **Factory Pattern**: EntityFactories create appropriate entity objects based on CAD entities
- **Command Pattern**: Operations encapsulated as Action delegates for flexible execution

## Usage

### Basic Usage Pattern

1. Create an orchestrator for the desired CAD application
2. Define operations to perform on the CAD document
3. Execute operations within the STA thread context

### AutoCAD Example

```csharp
// Create an orchestrator for AutoCAD operations
var orchestrator = new OrchestrateAutoCADExportDocumentData("path/to/drawing.dwg");

// Export all entities to JSON
orchestrator.ExportEntitiesToJson();

// Report on all layers in the document
orchestrator.ReportAllLayers();
```

### MicroStation Example

```csharp
// Create an orchestrator for MicroStation operations
var orchestrator = new OrchestrateMicroStationExportDocumentData("path/to/design.dgn");

// Export all entities to JSON
orchestrator.ExportEntitiesToJson();
```

## Important Implementation Details

- All CAD operations must run in a Single-Threaded Apartment (STA) thread due to COM interop requirements
- The library provides proper cleanup and disposal of COM objects to prevent memory leaks
- Error handling is implemented to catch and report COM exceptions

## Using the Test Project

The TestCADInterOpServices project provides examples of how to use the library and serves as entry points for testing functionality.

### Running Tests

1. Ensure sample CAD files are available in the `samples` directory
2. Run tests with the NUnit test runner
3. Tests must be run in STA mode (all test classes are decorated with `[Apartment(ApartmentState.STA)]`)

### Key Test Classes

- **AutoCADInterOpTests**: Examples for working with AutoCAD files
- **MicrostationInterOpTests**: Examples for working with MicroStation files
- **TestAnnotation**: Examples for working with annotations

## Common Operations

### Exporting CAD Data

```csharp
// Export to CSV
orchestrator.ExportEntitiesToCsv();

// Export to JSON
orchestrator.ExportEntitiesToJson();
```

### Working with Layouts and Layers

```csharp
// Report all layers
orchestrator.ReportAllLayers();

// List all layouts
orchestrator.ReportLayouts();
```

### Analyzing Block References

```csharp
// Report on block references and their components
orchestrator.ReportAllReferenceBlockComponents();
```

## Thread Safety

When using this library, be sure to run all CAD operations within a Single-Threaded Apartment (STA) thread. This is required for COM interop to function correctly.

Example:

```csharp
// Ensure the thread is STA
if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
{
    throw new Exception("The current thread must be set to STA state.");
}

// Proceed with CAD operations
```

## License

See the LICENSE file for details.