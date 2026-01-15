# Claude Code 项目指南

本文档为 AI 助手提供项目开发规范和协作指南。

## 项目概述

`linch-docx-rs` 是一个 Rust DOCX 读写库，核心特性是 **round-trip preservation**（未知元素保持不变）。

## 代码规范

### Rust 风格

- 所有代码必须通过 `cargo clippy -- -D warnings`
- 所有代码必须通过 `cargo fmt --check`
- 不使用 `#[allow(dead_code)]`，未使用的代码应删除
- 优先使用标准库，避免不必要的依赖

### 文件组织

- 单个文件不超过 800 行
- 超过则拆分为子模块
- 模块结构：`mod.rs` 只做导出，逻辑放在独立文件

### 错误处理

- 使用 `thiserror` 定义错误类型
- 提供有意义的错误信息
- 错误类型定义在 `src/error.rs`

## Git 工作流

### 分支策略

- `main` - 稳定分支，CI 必须通过
- 功能开发在 feature 分支
- PR 合并到 main

### Commit 规范 (Conventional Commits)

```
feat: 新功能 (minor 版本升级)
fix: bug 修复 (patch 版本升级)
feat!: 破坏性变更 (major 版本升级)
docs: 文档更新
ci: CI/CD 变更
refactor: 重构
test: 测试相关
chore: 杂项
```

### 发布流程

使用 `release-plz` 自动化：
1. 合并 PR 到 main
2. release-plz 根据 commits 计算版本号
3. 自动创建 Release PR
4. 合并后发布到 crates.io

## 测试要求

- 新功能必须有测试
- 测试文件放在 `tests/` 目录
- 单元测试放在对应模块的 `#[cfg(test)]` 块
- 运行全部测试：`cargo test`

## 架构说明

```
src/
├── lib.rs          # 库入口，导出公共 API
├── error.rs        # 错误类型
├── document/       # 文档层 (Paragraph, Run, Table)
├── opc/            # OPC 层 (Package, Part, Relationships)
└── xml/            # XML 层 (RawXmlNode, namespace)
```

### 核心设计原则

1. **Round-trip preservation**: 所有结构体都有 `unknown_children: Vec<RawXmlNode>`
2. **Lazy parsing**: 只解析需要的内容
3. **类型安全**: 充分利用 Rust 类型系统

## CI/CD

- `.github/workflows/ci.yml` - 测试、格式化、clippy
- `.github/workflows/release.yml` - 自动发布

## 常用命令

```bash
# 开发
cargo build
cargo test
cargo clippy -- -D warnings
cargo fmt

# 查看文档
cargo doc --open
```

## 注意事项

- 不要直接 push 到 main（应该 PR）
- 不要 ignore clippy 警告，要修复
- 不要添加未使用的依赖或代码
- commit message 要遵循规范
