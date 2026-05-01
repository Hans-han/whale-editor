# Task Plan: Office Agent Plugin MVP

## Goal
Build a TypeScript/Node.js Office document editing agent that integrates with DeepSeek V4, supporting .docx/.pptx files with structured patch operations, validation loops, and cache-friendly prompt building.

## Phases

### Phase 1: Project Setup
- [ ] Create package.json with dependencies
- [ ] Create tsconfig.json
- [ ] Create vitest.config.ts
- [ ] Create .env.example
- [ ] Create directory structure

### Phase 2: Core Types & Schemas
- [ ] Define TypeScript types for patches, manifests, operations
- [ ] Create Zod schemas for validation
- [ ] Create JSON schema files

### Phase 3: LLM Provider Adapter
- [ ] Define provider interface
- [ ] Implement DeepSeek provider
- [ ] Implement usage tracking

### Phase 4: Office File Handling
- [ ] Implement OOXML package handler (JSZip wrapper)
- [ ] Implement DOCX manifest generator
- [ ] Implement PPTX manifest generator
- [ ] Implement validation utilities

### Phase 5: Patch System
- [ ] Implement patch validator
- [ ] Implement DOCX patch executor
- [ ] Implement PPTX patch executor

### Phase 6: Agent Loop & Prompt Builder
- [ ] Create system prompt
- [ ] Implement cache-friendly prompt builder
- [ ] Implement agent loop with retry logic

### Phase 7: API Server
- [ ] Create Express/Fastify server
- [ ] Implement file upload routes
- [ ] Implement agent execution endpoint

### Phase 8: Tests
- [ ] Unit tests for manifest generation
- [ ] Unit tests for patch operations
- [ ] Unit tests for prompt builder
- [ ] Integration tests

### Phase 9: Documentation
- [ ] Create comprehensive README.md

## Key Decisions
- TypeScript with ESM modules
- JSZip for ZIP handling
- fast-xml-parser for XML parsing
- Zod for schema validation
- Express for API server
- Vitest for testing

## Status
**Currently in Phase 1** - Setting up project structure
