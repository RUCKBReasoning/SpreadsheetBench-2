sweagent run \
  --config config/visualisation.yaml \
  --env.deployment.image spreadsheetbench-v2 \
  --agent.model.name='openrouter/z-ai/glm-5' \
  --agent.model.api_key='' \
  --agent.model.completion_kwargs='{"extra_body": {"reasoning": {"enabled": true}}}' \
  --dataset_path ../data/Visualization \
