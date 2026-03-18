param(
  [int]$MaxRounds = 20,
  [string]$Model = "gpt-5",
  [switch]$AutoCommit,
  [switch]$Dangerous
)

$project = "C:\Users\LIN\ledger-analyzer"
$goal = Join-Path $project ".codex-autopilot-goal.md"
$runner = "C:\Users\LIN\.codex\skills\codex-batch-autopilot\scripts\run-batch-autopilot.ps1"

if (!(Test-Path $runner)) { throw "Runner not found: $runner" }
if (!(Test-Path $goal)) { throw "Goal file not found: $goal" }

$args = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $runner,
  "-ProjectDir", $project,
  "-GoalFile", $goal,
  "-MaxRounds", $MaxRounds,
  "-Model", $Model
)
if ($AutoCommit) { $args += "-AutoCommit" }
if ($Dangerous) { $args += "-Dangerous" }

& powershell @args
exit $LASTEXITCODE
