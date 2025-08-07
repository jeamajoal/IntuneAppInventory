<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# IntuneInventory PowerShell Module Development Instructions

This is a PowerShell module project for inventorying Microsoft Intune applications, scripts, and remediations with SQLite backend storage.

## Project Guidelines

- Follow PowerShell best practices and coding standards
- Use approved verbs for PowerShell functions (Get-, Set-, New-, Remove-, etc.)
- Implement proper error handling with try/catch blocks
- Use Write-Verbose for debugging information
- Include parameter validation and help documentation
- Follow the module structure with Private and Public function directories

## Module Architecture

- **Database**: SQLite for local storage and querying
- **Authentication**: Microsoft Graph API with proper scoping
- **Core Functions**: Inventory collection, reporting, and source code management
- **Data Structure**: Maintain both raw backend data and human-readable structured outputs

## Key Features to Implement

1. Intune API integration for apps, scripts, and remediations
2. SQLite database operations for data persistence
3. Assignment information collection and storage
4. Source code management with post-inventory addition capability
5. Comprehensive reporting and export functionality
6. Logging and authentication mechanisms (to be provided by user)

## Code Quality Standards

- Include comprehensive parameter help and examples
- Implement proper pipeline support where applicable
- Use consistent naming conventions
- Include unit tests for critical functions
- Ensure proper resource cleanup and connection management
