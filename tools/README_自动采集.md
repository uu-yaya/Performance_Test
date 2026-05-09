# 桌宠性能全指标自动采集

这套脚本用于把桌宠类产品的性能测试尽量自动化，覆盖 Steam、GitHub Release、GitHub 源码运行、官网安装包和自研 Demo。

## 1. 能自动采集的指标

`collect_desktop_pet_perf.ps1` 可直接采集：

- CPU：平均值、P95、最大值
- GPU：平均值、P95、最大值、GPU Engine 类型
- 内存：工作集内存、私有内存
- 显存：Dedicated VRAM、Shared VRAM
- 稳定性：句柄数、线程数、增长量
- IO：进程读写 KB/s
- 网络：TCP 连接数量
- 进程状态：进程数量、PID 列表、时间戳

这些指标会输出到：

- `perf_runs/<时间_产品_场景>/raw_samples.csv`
- `perf_runs/<时间_产品_场景>/summary.csv`
- `perf_runs/<时间_产品_场景>/metadata.json`

## 2. 需要外部工具日志的指标

以下指标需要专用工具导出 CSV 后再导入 Excel：

- FPS、1% Low、帧时间 P95：PresentMon 或 NVIDIA FrameView
- GPU Busy：PresentMon
- 功耗、温度、风扇：HWiNFO 传感器日志

原因是 Windows 普通进程 API 无法稳定提供每个应用的 FPS、1% Low、真实功耗和温度。

## 3. 单场景采集命令

```powershell
powershell -ExecutionPolicy Bypass -File D:\desktop_pet\tools\collect_desktop_pet_perf.ps1 `
  -ProductId P001 `
  -ProductName BongoCat `
  -ProcessName BongoCat `
  -Scenario 空闲常驻 `
  -DurationSec 600 `
  -IntervalSec 1
```

如果产品有多个进程：

```powershell
-ProcessName BongoCat,python,node
```

## 4. 按配置批量采集

复制并修改：

```text
D:\desktop_pet\tools\perf_collect_config.example.json
```

然后运行：

```powershell
powershell -ExecutionPolicy Bypass -File D:\desktop_pet\tools\run_perf_from_config.ps1 `
  -ConfigPath D:\desktop_pet\tools\perf_collect_config.example.json
```

## 5. 导入 Excel

采集结束后运行：

```powershell
& C:\Users\zoeuuliu\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe `
  D:\desktop_pet\tools\import_perf_results.py
```

如果有 PresentMon 和 HWiNFO 日志：

```powershell
& C:\Users\zoeuuliu\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe `
  D:\desktop_pet\tools\import_perf_results.py `
  --presentmon-csv D:\desktop_pet\external_logs\presentmon_bongocat.csv `
  --hwinfo-csv D:\desktop_pet\external_logs\hwinfo_bongocat.csv
```

结果会追加到：

- `桌宠类产品性能测试记录表.xlsx` 的 `场景记录`
- `桌宠类产品性能测试记录表.xlsx` 的 `采样数据`
- `桌宠类产品性能测试记录表.xlsx` 的 `自动采集导入`

## 6. 建议测试顺序

1. 系统基线：5 分钟
2. 启动峰值：2 分钟
3. 空闲常驻：10 分钟
4. 基础交互：5 分钟
5. 功能开启：10 分钟
6. 长时间稳定：60 分钟或更久
7. 游戏影响：开/关桌宠各跑一轮 PresentMon 或 FrameView

## 7. 注意事项

- 进程名不要带路径，`.exe` 可带可不带。
- GitHub 源码运行时要把 `python`、`node`、`electron` 等子进程一起写进 `ProcessName`。
- 采样期间尽量不要打开浏览器、下载器、录屏软件等干扰项。
- 长时间稳定性测试重点看内存、句柄、线程是否持续增长。
