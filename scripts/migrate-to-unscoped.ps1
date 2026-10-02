#!/usr/bin/env pwsh
<#
.SYNOPSIS
One-off move from @christophervr/ooxml-core and @christophervr/office-ui to the unscoped
ooxml-core and ooxml-ui, across this repository, docx-viewer and pptx-viewer.

.DESCRIPTION
Run it yourself in PowerShell 7: every npm and gh command is attached to your console, so npm's
browser login and confirmation prompts work as usual (no one-time codes are passed in). Each step
is safe to re-run; resume a failed run with -From <step>.

  preflight  tools, logins and the three checkouts; changes nothing
  pause      disable the release workflow in all three repositories
  unpublish  remove every scoped version that depends on the old names, dependents first;
             when npm refuses one, deprecate it instead and keep going
  tags       delete the git tags and GitHub releases of every version that is now gone
  ooxml      build, publish ooxml-core then ooxml-ui, refresh the lockfile, push main and tags
  docx       publish the seven unscoped docx packages from docx-viewer's main and tag them
  pptx       point pptx-viewer at ooxml-core, push, and dispatch a pptx-viewer release (4.9.4)

npm only allows these unpublishes while each package is under 72 hours old and nothing else
depends on it, so run before the deadline below. Trusted publishers can only be added to a
package that exists: the script stops after each first publish so you can add them on npmjs.com.

.EXAMPLE
./scripts/migrate-to-unscoped.ps1 -Pptx ../pptx-viewer-new -Docx ../docx-viewer -DryRun
./scripts/migrate-to-unscoped.ps1 -Pptx ../pptx-viewer-new -Docx ../docx-viewer
./scripts/migrate-to-unscoped.ps1 -Pptx ../pptx-viewer-new -Docx ../docx-viewer -From ooxml
#>
param(
	[string]$Pptx = (Join-Path $PSScriptRoot '../../pptx-viewer'),
	[string]$Docx = (Join-Path $PSScriptRoot '../../docx-viewer'),
	[ValidateSet('preflight', 'pause', 'unpublish', 'tags', 'ooxml', 'docx', 'pptx')]
	[string]$From = 'preflight',
	[ValidateSet('', 'preflight', 'pause', 'unpublish', 'tags', 'ooxml', 'docx', 'pptx')]
	[string]$Only = '',
	[switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandArgumentPassing = 'Standard'

$Ooxml = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Pptx = (Resolve-Path $Pptx).Path
$Docx = (Resolve-Path $Docx).Path
$Repos = @{ ooxml = 'ChristopherVR/ooxml'; docx = 'ChristopherVR/docx-viewer'; pptx = 'ChristopherVR/pptx-viewer' }
$Deadline = [datetime]::Parse('2026-10-04T13:39:00Z').ToUniversalTime()
$Trailer = 'Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>'
$Frameworks = 'react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'
$Steps = 'preflight', 'pause', 'unpublish', 'tags', 'ooxml', 'docx', 'pptx'

# Dependents first: npm refuses to unpublish a package another package depends on.
$Unpublish = @($Frameworks | ForEach-Object { "@christophervr/docx-$_-viewer" }) + @(
	'@christophervr/docx-core', 'pptx-viewer-core@4.9.3', '@christophervr/office-ui', '@christophervr/ooxml-core')
$Deprecation = @{
	'@christophervr/ooxml-core' = 'Renamed to ooxml-core. Install ooxml-core instead.'
	'@christophervr/office-ui'  = 'Renamed to ooxml-ui. Install ooxml-ui instead.'
	'@christophervr/docx-core'  = 'Renamed to docx-core. Install docx-core instead.'
	'pptx-viewer-core@4.9.3'    = 'Depended on the renamed @christophervr/ooxml-core. Use 4.9.4 or later.'
}
foreach ($f in $Frameworks) { $Deprecation["@christophervr/docx-$f-viewer"] = "Renamed to docx-$f-viewer. Install docx-$f-viewer instead." }

# A command that changes something: printed under -DryRun, otherwise run attached to the console.
function Invoke-Change([string]$Cwd, [string]$Exe, [string[]]$Arguments, [switch]$AllowFail) {
	Write-Host "  ($Cwd) $Exe $($Arguments -join ' ')" -ForegroundColor Cyan
	if ($DryRun) { return $true }
	Push-Location $Cwd
	try { & $Exe @Arguments; $ok = $LASTEXITCODE -eq 0 } finally { Pop-Location }
	if (-not $ok -and -not $AllowFail) { throw "Failed: $Exe $($Arguments -join ' ')" }
	return $ok
}

# A read-only command: always runs, returns its trimmed output ('' on failure).
function Invoke-Read([string]$Cwd, [string]$Exe, [string[]]$Arguments) {
	Push-Location $Cwd
	try { $out = & $Exe @Arguments 2>$null; if ($LASTEXITCODE -ne 0) { return '' } } finally { Pop-Location }
	return (($out | Out-String).Trim())
}

function Test-OnNpm([string]$Spec) { return [bool](Invoke-Read $Ooxml 'npm' @('view', $Spec, 'version')) }
function Get-Json([string]$Path) { return Get-Content -Raw $Path | ConvertFrom-Json }

function Wait-User([string]$Message) {
	Write-Host "`n>>> $Message" -ForegroundColor Yellow
	if (-not $DryRun) { [void](Read-Host 'Press Enter when done (Ctrl+C to stop; resume later with -From)') }
}

function Get-TrustedPublisherNote([string[]]$Packages, [string]$Repo) {
	return "Add a trusted publisher on npmjs.com for: $($Packages -join ', ')`n" +
	"    Package > Settings > Trusted Publisher > GitHub Actions: owner ChristopherVR,`n" +
	"    repository $($Repo.Split('/')[1]), workflow release.yml, environment npm"
}

function Assert-Clean([string]$Dir, [string]$Name) {
	$dirty = Invoke-Read $Dir 'git' @('status', '--porcelain', '--untracked-files=no')
	if ($dirty) { throw "$Name ($Dir) has uncommitted changes:`n$dirty" }
	$branch = Invoke-Read $Dir 'git' @('branch', '--show-current')
	if ($branch -ne 'main') { throw "$Name ($Dir) is on '$branch', not main." }
}

function Confirm-NpmLogin {
	$user = Invoke-Read $Ooxml 'npm' @('whoami')
	if (-not $user) {
		if ($DryRun) { Write-Warning 'Not logged in to npm (the real run will ask you to log in).'; return }
		Write-Host 'Logging in to npm (follow the browser prompt).'
		& npm login
		$user = Invoke-Read $Ooxml 'npm' @('whoami')
		if (-not $user) { throw 'npm login did not complete.' }
	}
	Write-Host "npm user: $user"
}

function Step-Preflight {
	if ([datetime]::UtcNow -gt $Deadline) { Write-Warning "Past $Deadline UTC: npm will likely refuse the unpublishes." }
	Confirm-NpmLogin
	& gh auth status; if ($LASTEXITCODE -ne 0) { throw 'Run gh auth login first.' }
	foreach ($pair in @(@($Ooxml, 'ooxml'), @($Docx, 'docx-viewer'), @($Pptx, 'pptx-viewer'))) {
		if (-not (Test-Path (Join-Path $pair[0] '.git'))) { throw "$($pair[1]) checkout not found at $($pair[0])" }
	}
	if ((Get-Json (Join-Path $Ooxml 'package.json')).name -ne 'ooxml-core') { throw 'This checkout does not have the ooxml-core rename committed.' }
	Assert-Clean $Ooxml 'ooxml'
	[void](Invoke-Read $Ooxml 'git' @('fetch', 'origin'))
	# The commit published to npm must reach main unchanged, so main may only fast-forward.
	Push-Location $Ooxml; & git merge-base --is-ancestor origin/main HEAD; $ff = $LASTEXITCODE -eq 0; Pop-Location
	if (-not $ff) { throw 'ooxml: origin/main has commits this checkout lacks; pull first.' }
	[void](Invoke-Read $Docx 'git' @('fetch', 'origin'))
	$docxCore = Invoke-Read $Docx 'git' @('show', 'origin/main:packages/core/package.json') | ConvertFrom-Json
	if ($docxCore.name -ne 'docx-core' -or -not $docxCore.dependencies.'ooxml-core') {
		throw 'docx-viewer origin/main is not on docx-core + ooxml-core yet: push that first.'
	}
	Write-Host 'Preflight passed.' -ForegroundColor Green
}

function Step-Pause {
	foreach ($repo in $Repos.Values) { [void](Invoke-Change $Ooxml 'gh' @('workflow', 'disable', 'release.yml', '-R', $repo) -AllowFail) }
}

function Step-Unpublish {
	Confirm-NpmLogin
	foreach ($spec in $Unpublish) {
		if (-not (Test-OnNpm $spec)) { Write-Host "${spec}: already gone"; continue }
		$whole = -not $spec.Substring(1).Contains('@')
		$unpublishArgs = @('unpublish', $spec) + $(if ($whole) { @('--force') } else { @() })
		if (-not (Invoke-Change $Ooxml 'npm' $unpublishArgs -AllowFail)) {
			Write-Warning "${spec}: npm refused the unpublish; deprecating instead."
			[void](Invoke-Change $Ooxml 'npm' @('deprecate', $spec, $Deprecation[$spec]) -AllowFail)
		}
	}
}

function Step-Tags {
	$docxTags = (Invoke-Read $Docx 'git' @('ls-remote', '--tags', '--refs', 'origin', 'refs/tags/@christophervr/docx-*')) -split "`n" |
		Where-Object { $_ } | ForEach-Object { ($_ -split 'refs/tags/')[1] }
	$ooxmlTags = (Invoke-Read $Ooxml 'git' @('tag', '-l', '@christophervr/*')) -split "`n" | Where-Object { $_ }
	$sets = @(@($Ooxml, $Repos.ooxml, $ooxmlTags), @($Docx, $Repos.docx, $docxTags), @($Pptx, $Repos.pptx, @('pptx-viewer-core@4.9.3')))
	foreach ($set in $sets) {
		foreach ($tag in $set[2]) {
			if (-not $tag) { continue }
			if (Test-OnNpm $tag) { Write-Host "${tag}: still on npm (not unpublished), keeping the tag"; continue }
			$hasRelease = [bool](Invoke-Read $set[0] 'gh' @('release', 'view', $tag, '-R', $set[1], '--json', 'tagName'))
			if ($hasRelease) { [void](Invoke-Change $set[0] 'gh' @('release', 'delete', $tag, '-R', $set[1], '--cleanup-tag', '-y')) }
			else { [void](Invoke-Change $set[0] 'git' @('push', 'origin', ":refs/tags/$tag") -AllowFail) }
			[void](Invoke-Change $set[0] 'git' @('tag', '-d', $tag) -AllowFail)
		}
	}
}

function Step-Ooxml {
	Confirm-NpmLogin
	$core = (Get-Json (Join-Path $Ooxml 'package.json')).version
	$ui = (Get-Json (Join-Path $Ooxml 'packages/ui/package.json')).version
	$published = Invoke-Read $Ooxml 'git' @('rev-parse', 'HEAD')
	[void](Invoke-Change $Ooxml 'bun' @('run', 'build'))
	[void](Invoke-Change $Ooxml 'bun' @('run', '--cwd', 'packages/ui', 'build'))
	[void](Invoke-Change $Ooxml 'node' @('scripts/publish-released.mjs', '--tag', "ooxml-core@$core", '--manual'))
	[void](Invoke-Change $Ooxml 'node' @('scripts/publish-released.mjs', '--tag', "ooxml-ui@$ui", '--manual'))
	# The UI workspace resolves ooxml-core from the registry, so the lockfile changes now.
	[void](Invoke-Change $Ooxml 'bun' @('install'))
	if ($DryRun -or (Invoke-Read $Ooxml 'git' @('status', '--porcelain', 'bun.lock'))) {
		[void](Invoke-Change $Ooxml 'git' @('add', 'bun.lock'))
		[void](Invoke-Change $Ooxml 'git' @('commit', '-m', 'chore(deps): lock the unscoped ooxml-core in the UI workspace', '-m', $Trailer))
	}
	Wait-User (Get-TrustedPublisherNote @('ooxml-core', 'ooxml-ui') $Repos.ooxml)
	foreach ($tag in "ooxml-core@$core", "ooxml-ui@$ui") { [void](Invoke-Change $Ooxml 'git' @('tag', $tag, $published) -AllowFail) }
	[void](Invoke-Change $Ooxml 'git' @('push', 'origin', 'HEAD:main'))
	[void](Invoke-Change $Ooxml 'git' @('push', 'origin', "ooxml-core@$core", "ooxml-ui@$ui"))
	[void](Invoke-Change $Ooxml 'gh' @('workflow', 'enable', 'release.yml', '-R', $Repos.ooxml))
}

function Step-Docx {
	Confirm-NpmLogin
	Assert-Clean $Docx 'docx-viewer'
	[void](Invoke-Change $Docx 'git' @('pull', '--ff-only', 'origin', 'main'))
	$json = Invoke-Read $Docx 'node' @('-e', "import('./scripts/release-plan.mjs').then(m => console.log(JSON.stringify(Object.values(m.PACKAGES))))")
	# docx-core first: every framework package depends on it.
	$targets = @($json | ConvertFrom-Json | Sort-Object { if ($_.npm -eq 'docx-core') { 0 } else { 1 } } |
		ForEach-Object { [pscustomobject]@{ npm = $_.npm; dir = $_.dir; version = (Get-Json (Join-Path $Docx "$($_.dir)/package.json")).version } })
	[void](Invoke-Change $Docx 'bun' @('install', '--frozen-lockfile'))
	[void](Invoke-Change $Docx 'bun' @('run', 'build:packages'))
	[void](Invoke-Change $Docx 'bun' @('run', 'check:published'))
	$published = Invoke-Read $Docx 'git' @('rev-parse', 'HEAD')
	foreach ($t in $targets) {
		if (Test-OnNpm "$($t.npm)@$($t.version)") { Write-Host "$($t.npm)@$($t.version): already on npm" }
		else { [void](Invoke-Change (Join-Path $Docx $t.dir) 'npm' @('publish', '--access', 'public')) }
		[void](Invoke-Change $Docx 'git' @('tag', "$($t.npm)@$($t.version)", $published) -AllowFail)
	}
	[void](Invoke-Change $Docx 'git' (@('push', 'origin') + @($targets | ForEach-Object { "$($_.npm)@$($_.version)" })))
	Wait-User (Get-TrustedPublisherNote @($targets.npm) $Repos.docx)
	[void](Invoke-Change $Docx 'gh' @('workflow', 'enable', 'release.yml', '-R', $Repos.docx))
}

function Step-Pptx {
	Assert-Clean $Pptx 'pptx-viewer'
	[void](Invoke-Change $Pptx 'git' @('pull', '--ff-only', 'origin', 'main'))
	$coreVersion = (Get-Json (Join-Path $Ooxml 'package.json')).version
	$files = (Invoke-Read $Pptx 'git' @('grep', '-l', '-I', '@christophervr/ooxml-core', '--', ':!*CHANGELOG.md', ':!bun.lock')) -split "`n" | Where-Object { $_ }
	foreach ($file in $files) {
		$path = Join-Path $Pptx $file
		$text = (Get-Content -Raw $path).Replace('@christophervr/ooxml-core', 'ooxml-core')
		# The old ^0.1.0 range cannot reach 0.2.x; use the version just published.
		if ($file -eq 'packages/core/package.json') { $text = $text -replace '"ooxml-core": "[^"]*"', "`"ooxml-core`": `"^$coreVersion`"" }
		Write-Host "  rewrite $file" -ForegroundColor Cyan
		if (-not $DryRun) { [IO.File]::WriteAllText($path, $text) }
	}
	[void](Invoke-Change $Pptx 'bun' @('install'))
	[void](Invoke-Change $Pptx 'bun' @('run', '--filter', 'pptx-viewer-core', 'build'))
	[void](Invoke-Change $Pptx 'git' (@('add') + $files + @('bun.lock')))
	[void](Invoke-Change $Pptx 'git' @('commit', '-m', 'build(core): depend on the unscoped ooxml-core package', '-m',
			'@christophervr/ooxml-core is now published as ooxml-core. pptx-viewer-core@4.9.3, the one version that depended on the scoped name, was unpublished; this releases 4.9.4.',
			'-m', $Trailer))
	[void](Invoke-Change $Pptx 'git' @('push', 'origin', 'HEAD:main'))
	[void](Invoke-Change $Pptx 'gh' @('workflow', 'enable', 'release.yml', '-R', $Repos.pptx))
	[void](Invoke-Change $Pptx 'gh' @('workflow', 'run', 'release.yml', '-R', $Repos.pptx, '--ref', 'main'))
}

$selected = if ($Only) { @($Only) } else { $Steps[$Steps.IndexOf($From)..($Steps.Count - 1)] }
if ($selected -notcontains 'preflight') { Step-Preflight }
foreach ($step in $selected) {
	Write-Host "`n=== $step ===" -ForegroundColor Magenta
	& "Step-$((Get-Culture).TextInfo.ToTitleCase($step))"
}
Write-Host "`nDone. Watch the pptx-viewer release: gh run list -R ChristopherVR/pptx-viewer -w release.yml" -ForegroundColor Green
