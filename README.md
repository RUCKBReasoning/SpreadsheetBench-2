# SpreadsheetBench 2

A benchmark for evaluating AI agents on spreadsheet manipulation and analysis tasks, including financial modeling, data visualization, formula debugging, and template-based operations. Built on the [SWE-agent](https://github.com/princeton-nlp/SWE-agent) framework.

## Requirements

- Python 3.11
- Conda
- Docker

## Installation

### 1. Set up the Conda environment

```bash
conda create -n ssb-v2 python==3.11 -y
conda activate ssb-v2

cd SWE-agent
pip install --upgrade pip && pip install --editable .
```

### 2. Build the Docker image

```bash
docker build -f spreadsheet.Dockerfile -t spreadsheetbench-v2 .
```

## Dataset

Place the dataset under the `data/` directory. The benchmark includes four task categories:

| Category        | Description                             |
| --------------- | --------------------------------------- |
| Debugging       | Formula debugging and error correction  |
| Financial_Model | Financial modeling and calculation      |
| Template        | Template-based spreadsheet operations   |
| Visualization   | Chart generation and data visualization |

Each category folder should contain a `dataset.json` file and the corresponding spreadsheet files.

## Running Experiments

Run SWE-agent from the `SWE-agent/` directory. A complete runnable example is
provided in `SWE-agent/scripts/example.sh`.

```bash
conda activate ssb-v2
cd SWE-agent

sweagent run \
  --config config/spreadsheet.yaml \
  --env.deployment.image spreadsheetbench-v2 \
  --agent.model.name='openrouter/z-ai/glm-5' \
  --agent.model.api_key='<your_api_key>' \
  --agent.model.completion_kwargs='{"extra_body": {"reasoning": {"enabled": true}}}' \
  --dataset_path ../data/<Category>
```

Replace `<Category>` with one of `Debugging`, `Financial_Model`, `Template`, or
`Visualization`. For `Visualization` tasks, you must use
`config/visualisation.yaml` instead of `config/spreadsheet.yaml`. For example,
to run the visualization split:

```bash
bash scripts/example.sh
```

## Evaluation

After obtaining model outputs, first refresh the cached spreadsheet values with
LibreOffice:

```bash
python evaluation/open_spreadsheet.py \
  --dir_path <path_to_output_excel>
```

Then run the task-specific evaluator.

For `Debugging`, `Financial_Model`, and `Template` tasks:

```bash
python evaluation/evaluation.py \
  --model <model_name> \
  --dataset <Category> \
  --outputs-dir <path_to_output_excel> \
  --workers <N>
```

Replace `<Category>` with `Debugging`, `Financial_Model`, or `Template`. The
evaluator reports regression accuracy, modification accuracy, and overall
accuracy, and writes results to `results/<Category>/`.

For `Visualization` tasks, use the VLM checklist evaluator with `glm-4.6v`:

```bash
python evaluation/run_visual_vlm_checklist_eval.py \
  --tasks-json data/Visualization/dataset.json \
  --output-dir <path_to_output_excel> \
  --api-key <your_bigmodel_api_key> \
  --model glm-4.6v
```

You can also provide the VLM API key through the environment:

```bash
export VLM_API_KEY=<your_bigmodel_api_key>
python evaluation/run_visual_vlm_checklist_eval.py \
  --tasks-json data/Visualization/dataset.json \
  --output-dir <path_to_output_excel> \
  --model glm-4.6v
```

The visualization evaluator saves a JSON report next to the output directory by
default, named `evaluation_report_<output-dir-name>.json`.

## Project Structure

```
├── data/                        # Benchmark datasets
│   ├── Debugging/
│   ├── Financial_Model/
│   ├── Template/
│   └── Visualization/
├── evaluation/
│   ├── evaluation.py            # Main evaluation script
│   ├── open_spreadsheet.py      # LibreOffice formula refresh utility
│   └── run_visual_vlm_checklist_eval.py  # Visualization VLM evaluator
├── SWE-agent/                   # SWE-agent framework (modified)
│   ├── config/
│   │   ├── spreadsheet.yaml     # Config for non-visualization tasks
│   │   └── visualisation.yaml   # Config for visualization tasks
│   └── ...
└── spreadsheet.Dockerfile       # Docker image definition
```
