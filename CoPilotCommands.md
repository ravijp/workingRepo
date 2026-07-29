The publish failed with PromptAgentDefinition.__init__() got an unexpected keyword argument 'description'. Fix it as follows, and do not guess at optional fields.

1. In foundry/deploy/publish_agent.py, _sdk_definition must take only (model, instructions, schema) and construct PromptAgentDefinition with only model, instructions, and text=PromptAgentDefinitionTextOptions(format=TextResponseFormatJsonSchema(name="CaseIntakeClassification", schema=schema, strict=True)). Remove the description argument entirely.

2. Idempotency must not depend on the description. Change _existing to take the instruction text as a fourth argument and recognise an already-published version by comparing that text against each existing version's instructions (read from version.definition.instructions, falling back to version.instructions), keeping the existing marker check as a first attempt. Update the caller to pass instructions.

3. Add a _create_version(client, agent_name, sdk_definition, description) helper that inspects inspect.signature(client.agents.create_version).parameters and passes description= if that parameter exists, else metadata={"fingerprint": description} if that exists, else neither. Catch TypeError around each attempt and fall through.

Before changing anything, run this and paste the output so we work from the real surface rather than assumptions:

uv run python -c "from azure.ai.projects.models import PromptAgentDefinition as P; import azure.ai.projects as m; print('SDK', m.__version__); print('fields', sorted(getattr(P,'_attribute_map',{}).keys()) or [k for k in P().__dict__]); import inspect; from azure.ai.projects.operations import AgentsOperations as A; print('create_version', inspect.signature(A.create_version))"

---
