# Excel batch 拆分与复原

把一个 Excel 工作表按固定数据行数拆成多个有序 batch，方便分批翻译、检查或
其他作业。每个 batch 都重复保留指定数量的表头行；处理完成后，可以按拆分时
生成的 manifest 将所有 batch 精确回填到原始行位置。

## GUI

打开 `qatools gui`，在【其他】中选择【Excel batch 拆分与复原】：

- 【拆分 batch】：选择 Excel、工作表、每批行数和表头行数。
- 【复原文件】：选择拆分时生成的完整 batch 目录。

默认每个 batch 包含 1000 个数据行，并保留第 1 行作为表头。

## CLI

拆分：

```bash
qatools batch split ./input.xlsx \
  --sheet Sheet1 \
  --batch-size 1000 \
  --header-rows 1 \
  --output-dir ./input_batches
```

复原：

```bash
qatools batch restore ./input_batches \
  --output ./input_restored.xlsx
```

`restore` 也可以直接接收 `batch_manifest.json` 的路径。

## 输出目录

拆分目录包含：

- `<原文件名>_batch_001_of_003.xlsx`：可分发作业的 batch 文件。
- `batch_manifest.json`：原始行范围和 batch 顺序。
- `_qatools_restore_source_<原文件名>.xlsx`：用于保留原工作簿结构的复原模板。

复原前请保留整个目录，不要修改或删除 manifest 和复原模板。工具会校验复原模板
的 SHA-256；模板发生变化时会停止复原，避免把 batch 写入错误底稿。

## 行为边界

- 支持 `.xlsx` 和 `.xlsm`；`.xlsm` 会保留 VBA 内容。
- `.xlsx` batch 只包含所选工作表，避免大文件的其他工作表在每个 batch 中重复
  占用时间和空间；复原时仍以完整原始模板为底稿，因此其他工作表不会丢失。
- `.xlsm` 为确保 VBA 内容完整，batch 会继续保留整本工作簿。
- 数据行从表头之后开始，直到工作表最后一个有值的单元格所在行；中间空行仍计入
  batch 行数。
- 复原会采用第一个 batch 的表头，并支持把 batch 中新增的结果列合并回完整文件。
- 复原回填单元格内容和公式；工作簿结构与单元格格式以复原模板为准。
- 数组公式和数据表公式需先在 batch 中转换为普通公式或值再复原。
- 如果 batch 在预期数据范围之后增加了新数据行，复原会报错，因为这些行没有可靠
  的原始位置。新增列不受此限制。
