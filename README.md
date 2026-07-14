# pacman / PacStudent Unity 项目快照

[中文](#中文) | [English](#english)

## 中文

这是一个名为 PacStudent 的 Unity 项目快照。仓库包含启动场景、Unity 项目设置、Package Manager 清单和生成的 C# 工程文件。

### 仓库中可以确认的内容

- Unity Editor 版本记录为 `2023.2.10f1c1`
- 默认场景为 `StartScene.unity`
- 场景中包含分数 UI、按钮、摄像机、灯光和 EventSystem
- Package 清单包含 2D Sprite、uGUI、Visual Scripting、Timeline 等 Unity 模块
- `Assembly-CSharp.csproj` 引用了 PacStudent、Ghost、Cherry、LevelGenerator 和音频等脚本

### 完整性说明

> 这不是一个完整、可直接构建的 Unity 工程。当前 GitHub 仓库没有包含 `Assembly-CSharp.csproj` 所引用的 `Assets/Script/*.cs`，也缺少大部分纹理与其他 Assets。根目录中还包含 Unity 导入日志和生成文件。

若要恢复可运行版本，需要从原始项目或备份中找回缺失的 `Assets` 内容，特别是：

- `PacStudentController.cs`
- `PacStudentMovement.cs`
- `GhostController.cs` 与 `Gost.cs`
- `CherryController.cs`
- `LevelGenerator.cs` 与 `LevelBuilder.cs`
- `StartScene.cs`、`BGAudio.cs` 与 `Shoot.cs`
- 场景引用的精灵、纹理、音频和 Prefab

### 尝试打开

1. 安装与 `2023.2.10f1c1` 兼容的 Unity Editor。
2. 使用 Unity Hub 选择仓库根目录。
3. 打开 `StartScene.unity`。
4. 如果 Unity 报告脚本或资源丢失，请先恢复上述文件，不要把缺失引用自动保存回场景。

### 建议的仓库整理

- 恢复完整的 `Assets/`、`Packages/` 和 `ProjectSettings/` 结构
- 添加标准 Unity `.gitignore`
- 从版本控制中移除 `Library`、`Logs`、导入日志、IDE 工程文件和 Shader 编译日志
- 在恢复后使用干净 clone 重新验证场景和构建

当前仓库没有附加许可证，因此默认保留全部权利。

## English

This repository is a snapshot of a Unity project named PacStudent. It contains a start scene, Unity project settings, a Package Manager manifest, and a generated C# project file.

### What can be verified in the repository

- The recorded Unity Editor version is `2023.2.10f1c1`
- The default scene is `StartScene.unity`
- The scene contains score UI, buttons, a camera, lighting, and an EventSystem
- The package manifest includes Unity modules such as 2D Sprite, uGUI, Visual Scripting, and Timeline
- `Assembly-CSharp.csproj` references PacStudent, ghost, cherry, level-generation, and audio scripts

### Completeness notice

> This is not a complete, directly buildable Unity project. The GitHub repository does not include the `Assets/Script/*.cs` files referenced by `Assembly-CSharp.csproj`, and most textures and other assets are also absent. The root additionally contains Unity import logs and generated files.

To restore a runnable version, recover the missing `Assets` content from the original project or a backup, especially:

- `PacStudentController.cs`
- `PacStudentMovement.cs`
- `GhostController.cs` and `Gost.cs`
- `CherryController.cs`
- `LevelGenerator.cs` and `LevelBuilder.cs`
- `StartScene.cs`, `BGAudio.cs`, and `Shoot.cs`
- Sprites, textures, audio, and prefabs referenced by the scene

### Attempt to open the snapshot

1. Install a Unity Editor version compatible with `2023.2.10f1c1`.
2. Select the repository root in Unity Hub.
3. Open `StartScene.unity`.
4. If Unity reports missing scripts or assets, restore them before saving the scene so missing references are not written back permanently.

### Recommended repository cleanup

- Restore the standard `Assets/`, `Packages/`, and `ProjectSettings/` structure
- Add a standard Unity `.gitignore`
- Remove `Library`, `Logs`, import logs, IDE project files, and shader compiler logs from version control
- Revalidate the scene and build from a clean clone after restoration

No license is included in the current repository, so all rights are reserved by default.

