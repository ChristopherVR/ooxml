#!/usr/bin/env pwsh
<#
.SYNOPSIS
One-off move from @christophervr/ooxml-core and @christophervr/office-ui to the unscoped
ooxml-core and ooxml-ui, across this repository, docx-viewer and pptx-viewer.

.DESCRIPTION
Run it yourself in PowerShell 7: every npm and gh command is attached to your console, and npm
uses web authentication (--auth-type=web), so it opens the browser to confirm instead of asking
for a one-time code. Each step
is safe to re-run. The script stops at the first failure; fix it and resume with -From <step>.

  preflight  tools, logins and the three checkouts; changes nothing
  pause      disable the release workflow in all three repositories
  unpublish  remove the scoped docx packages and @christophervr/office-ui, dependents first;
             deprecate pptx-viewer-core@4.9.3 and @christophervr/ooxml-core, which npm will not
             let go because pptx-viewer-core has dependents
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
	'@christophervr/docx-core', '@christophervr/office-ui')

# These cannot be unpublished, so they are deprecated instead. pptx-viewer-core has dependents
# (pptx-angular-viewer, pptx-viewer-mcp), and npm refuses to unpublish any version of such a
# package (E405). Its 4.9.3 depends on @christophervr/ooxml-core, which therefore stays too.
$Deprecate = [ordered]@{
	'pptx-viewer-core@4.9.3'    = 'Depends on @christophervr/ooxml-core, now published as ooxml-core. Use 4.9.4 or later.'
	'@christophervr/ooxml-core' = 'Renamed to ooxml-core. Install ooxml-core instead.'
}

# A command that changes something: printed under -DryRun, otherwise run with the console as its
# stdin/stdout so npm can show its browser login prompt. It must never be captured (no [void],
# no assignment): a captured native command loses the terminal and npm fails with EOTP.
# Any failure stops the script. Each step checks the current state first and skips what is
# already done.
function Invoke-Change([string]$Cwd, [string]$Exe, [string[]]$Arguments) {
	Write-Host "  ($Cwd) $Exe $($Arguments -join ' ')" -ForegroundColor Cyan
	if ($DryRun) { return }
	Push-Location $Cwd
	try { & $Exe @Arguments; $ok = $LASTEXITCODE -eq 0 } finally { Pop-Location }
	if (-not $ok) { throw "Stopped: '$Exe $($Arguments -join ' ')' failed (exit $LASTEXITCODE). Fix it and resume with -From." }
}

# npm commands that need the account: web authentication opens the browser instead of asking
# for a one-time code.
function Invoke-Npm([string]$Cwd, [string[]]$Arguments) { Invoke-Change $Cwd 'npm' $Arguments }
# Also reaches the npm that scripts/publish-released.mjs runs.
$env:npm_config_auth_type = 'web'

# A read-only command: always runs, returns its trimmed output ('' on failure).
function Invoke-Read([string]$Cwd, [string]$Exe, [string[]]$Arguments) {
	Push-Location $Cwd
	try { $out = & $Exe @Arguments 2>$null; if ($LASTEXITCODE -ne 0) { return '' } } finally { Pop-Location }
	return (($out | Out-String).Trim())
}

# --prefer-online skips npm's local cache, which still answers for a package just unpublished.
function Test-OnNpm([string]$Spec) { return [bool](Invoke-Read $Ooxml 'npm' @('view', $Spec, 'version', '--prefer-online')) }

# Native output on Windows ends lines with CRLF: split on both and trim, or every name keeps a \r.
function Split-Lines([string]$Text) { return @($Text -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

function Test-Deprecated([string]$Spec) { return [bool](Invoke-Read $Ooxml 'npm' @('view', $Spec, 'deprecated', '--prefer-online')) }
function Test-LocalTag([string]$Dir, [string]$Tag) { return [bool](Invoke-Read $Dir 'git' @('tag', '-l', $Tag)) }
function Test-RemoteTag([string]$Dir, [string]$Tag) {
	return [bool](Invoke-Read $Dir 'git' @('ls-remote', '--tags', '--refs', 'origin', "refs/tags/$Tag"))
}
function Get-RemoteTags([string]$Dir, [string]$Pattern) {
	return @(Split-Lines (Invoke-Read $Dir 'git' @('ls-remote', '--tags', '--refs', 'origin', "refs/tags/$Pattern")) |
		ForEach-Object { ($_ -split 'refs/tags/')[1] })
}
# Tag a published version at the commit npm recorded for it (gitHead), so a resumed run that
# has since committed the lockfile still tags the commit that was actually published.
function Add-Tag([string]$Dir, [string]$Tag) {
	if (Test-LocalTag $Dir $Tag) { Write-Host "${Tag}: tag already exists"; return }
	$sha = Invoke-Read $Ooxml 'npm' @('view', $Tag, 'gitHead', '--prefer-online')
	if (-not $sha) { $sha = Invoke-Read $Dir 'git' @('rev-parse', 'HEAD') }
	Invoke-Change $Dir 'git' @('tag', $Tag, $sha)
}

# Turn a repository's release workflow on or off, only when it is not already in that state.
function Set-ReleaseWorkflow([string]$Repo, [bool]$Enabled) {
	$state = Invoke-Read $Ooxml 'gh' @('api', "repos/$Repo/actions/workflows/release.yml", '--jq', '.state')
	$word = if ($Enabled) { 'enabled' } else { 'disabled' }
	if (($state -eq 'active') -eq $Enabled) { Write-Host "${Repo}: release workflow already $word"; return }
	Invoke-Change $Ooxml 'gh' @('workflow', $(if ($Enabled) { 'enable' } else { 'disable' }), 'release.yml', '-R', $Repo)
}

# The registry's CDN can serve a removed package for a short while after npm confirms the unpublish.
function Wait-GoneFromNpm([string]$Spec) {
	for ($i = 0; $i -lt 12; $i++) {
		if (-not (Test-OnNpm $Spec)) { return }
		Start-Sleep -Seconds 5
	}
	throw "Stopped: $Spec is still on npm a minute after the unpublish."
}

# The other way round: a fresh publish can take a moment to become visible.
function Wait-OnNpm([string]$Spec) {
	for ($i = 0; $i -lt 24; $i++) {
		if (Test-OnNpm $Spec) { return }
		Start-Sleep -Seconds 5
	}
	throw "Stopped: $Spec is not visible on npm two minutes after publishing."
}
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
	# Machine-local agent settings never get committed by this script, so they may be edited.
	$dirty = (Split-Lines (Invoke-Read $Dir 'git' @('status', '--porcelain', '--untracked-files=no')) |
		Where-Object { $_ -notmatch '\.claude/settings\.local\.json$' }) -join "`n"
	if ($dirty) { throw "$Name ($Dir) has uncommitted changes:`n$dirty" }
	$branch = Invoke-Read $Dir 'git' @('branch', '--show-current')
	if ($branch -ne 'main') { throw "$Name ($Dir) is on '$branch', not main." }
}

function Confirm-NpmLogin {
	$user = Invoke-Read $Ooxml 'npm' @('whoami')
	if (-not $user) {
		if ($DryRun) { Write-Warning 'Not logged in to npm (the real run will ask you to log in).'; return }
		Write-Host 'Logging in to npm (follow the browser prompt).'
		& npm login --auth-type=web
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
	foreach ($repo in $Repos.Values) { Set-ReleaseWorkflow $repo $false }
}

function Step-Unpublish {
	Confirm-NpmLogin
	foreach ($spec in $Unpublish) {
		if (-not (Test-OnNpm $spec)) { Write-Host "${spec}: already gone"; continue }
		$whole = -not $spec.Substring(1).Contains('@')
		Invoke-Npm $Ooxml (@('unpublish', $spec) + $(if ($whole) { @('--force') } else { @() }))
		if (-not $DryRun) { Wait-GoneFromNpm $spec }
	}
	foreach ($spec in $Deprecate.Keys) {
		if (Test-Deprecated $spec) { Write-Host "${spec}: already deprecated"; continue }
		Invoke-Npm $Ooxml @('deprecate', $spec, $Deprecate[$spec])
	}
}

function Step-Tags {
	$sets = @(
		@($Ooxml, $Repos.ooxml, (Get-RemoteTags $Ooxml '@christophervr/*')),
		@($Docx, $Repos.docx, (Get-RemoteTags $Docx '@christophervr/docx-*')),
		@($Pptx, $Repos.pptx, @('pptx-viewer-core@4.9.3')))
	foreach ($set in $sets) {
		$dir, $repo, $tags = $set
		foreach ($tag in $tags) {
			if (Test-OnNpm $tag) { Write-Host "${tag}: still on npm, keeping the tag"; continue }
			if (Invoke-Read $dir 'gh' @('release', 'view', $tag, '-R', $repo, '--json', 'tagName')) {
				Invoke-Change $dir 'gh' @('release', 'delete', $tag, '-R', $repo, '--cleanup-tag', '-y')
			}
			elseif (Test-RemoteTag $dir $tag) { Invoke-Change $dir 'git' @('push', 'origin', ":refs/tags/$tag") }
			else { Write-Host "${tag}: no tag or release on GitHub" }
			if (Test-LocalTag $dir $tag) { Invoke-Change $dir 'git' @('tag', '-d', $tag) }
		}
	}
}

function Step-Ooxml {
	Confirm-NpmLogin
	$core = (Get-Json (Join-Path $Ooxml 'package.json')).version
	$ui = (Get-Json (Join-Path $Ooxml 'packages/ui/package.json')).version
	if ((Test-OnNpm "ooxml-core@$core") -and (Test-OnNpm "ooxml-ui@$ui")) {
		Write-Host "ooxml-core@$core and ooxml-ui@${ui}: already on npm"
	}
	else {
		Invoke-Change $Ooxml 'bun' @('run', 'build')
		Invoke-Change $Ooxml 'bun' @('run', '--cwd', 'packages/ui', 'build')
		# publish-released.mjs skips a version that is already on npm.
		Invoke-Change $Ooxml 'node' @('scripts/publish-released.mjs', '--tag', "ooxml-core@$core", '--manual')
		# It also refuses the UI until the core it depends on is visible on npm.
		if (-not $DryRun) { Wait-OnNpm "ooxml-core@$core" }
		Invoke-Change $Ooxml 'node' @('scripts/publish-released.mjs', '--tag', "ooxml-ui@$ui", '--manual')
	}
	# The UI workspace resolves ooxml-core from the registry, so the lockfile changes now.
	Invoke-Change $Ooxml 'bun' @('install')
	if ($DryRun -or (Invoke-Read $Ooxml 'git' @('status', '--porcelain', 'bun.lock'))) {
		Invoke-Change $Ooxml 'git' @('add', 'bun.lock')
		Invoke-Change $Ooxml 'git' @('commit', '-m', 'chore(deps): lock the unscoped ooxml-core in the UI workspace', '-m', $Trailer)
	}
	Wait-User (Get-TrustedPublisherNote @('ooxml-core', 'ooxml-ui') $Repos.ooxml)
	foreach ($tag in "ooxml-core@$core", "ooxml-ui@$ui") { Add-Tag $Ooxml $tag }
	Invoke-Change $Ooxml 'git' @('push', 'origin', 'HEAD:main')
	Invoke-Change $Ooxml 'git' @('push', 'origin', "ooxml-core@$core", "ooxml-ui@$ui")
	Set-ReleaseWorkflow $Repos.ooxml $true
}

function Step-Docx {
	Confirm-NpmLogin
	Assert-Clean $Docx 'docx-viewer'
	Invoke-Change $Docx 'git' @('pull', '--ff-only', 'origin', 'main')
	$json = Invoke-Read $Docx 'node' @('-e', "import('./scripts/release-plan.mjs').then(m => console.log(JSON.stringify(Object.values(m.PACKAGES))))")
	# docx-core first: every framework package depends on it.
	$targets = @($json | ConvertFrom-Json | Sort-Object { if ($_.npm -eq 'docx-core') { 0 } else { 1 } } |
		ForEach-Object { [pscustomobject]@{ npm = $_.npm; dir = $_.dir; version = (Get-Json (Join-Path $Docx "$($_.dir)/package.json")).version } })
	$missing = @($targets | Where-Object { -not (Test-OnNpm "$($_.npm)@$($_.version)") })
	if ($missing.Count -gt 0) {
		Invoke-Change $Docx 'bun' @('install', '--frozen-lockfile')
		Invoke-Change $Docx 'bun' @('run', 'build:packages')
		Invoke-Change $Docx 'bun' @('run', 'check:published')
	}
	foreach ($t in $targets) {
		if (Test-OnNpm "$($t.npm)@$($t.version)") { Write-Host "$($t.npm)@$($t.version): already on npm" }
		else { Invoke-Npm (Join-Path $Docx $t.dir) @('publish', '--access', 'public') }
		Add-Tag $Docx "$($t.npm)@$($t.version)"
	}
	Invoke-Change $Docx 'git' (@('push', 'origin') + @($targets | ForEach-Object { "$($_.npm)@$($_.version)" }))
	Wait-User (Get-TrustedPublisherNote @($targets.npm) $Repos.docx)
	Set-ReleaseWorkflow $Repos.docx $true
}

function Step-Pptx {
	Assert-Clean $Pptx 'pptx-viewer'
	Invoke-Change $Pptx 'git' @('pull', '--ff-only', 'origin', 'main')
	$coreVersion = (Get-Json (Join-Path $Ooxml 'package.json')).version
	$files = Split-Lines (Invoke-Read $Pptx 'git' @('grep', '-l', '-I', '@christophervr/ooxml-core', '--', ':!*CHANGELOG.md', ':!bun.lock'))
	if ($files.Count -eq 0) {
		Write-Host 'pptx-viewer: already on ooxml-core'
		Set-ReleaseWorkflow $Repos.pptx $true
		$latest = Invoke-Read $Ooxml 'npm' @('view', 'pptx-viewer-core', 'version', '--prefer-online')
		if ([version]$latest -gt [version]'4.9.3') { Write-Host "pptx-viewer-core@${latest}: already released"; return }
		Invoke-Change $Pptx 'gh' @('workflow', 'run', 'release.yml', '-R', $Repos.pptx, '--ref', 'main')
		return
	}
	foreach ($file in $files) {
		$path = Join-Path $Pptx $file
		$text = (Get-Content -Raw $path).Replace('@christophervr/ooxml-core', 'ooxml-core')
		# The old ^0.1.0 range cannot reach 0.2.x; use the version just published.
		if ($file -eq 'packages/core/package.json') { $text = $text -replace '"ooxml-core": "[^"]*"', "`"ooxml-core`": `"^$coreVersion`"" }
		Write-Host "  rewrite $file" -ForegroundColor Cyan
		if (-not $DryRun) { [IO.File]::WriteAllText($path, $text) }
	}
	Invoke-Change $Pptx 'bun' @('install')
	Invoke-Change $Pptx 'bun' @('run', '--filter', 'pptx-viewer-core', 'build')
	Invoke-Change $Pptx 'git' (@('add') + $files + @('bun.lock'))
	Invoke-Change $Pptx 'git' @('commit', '-m', 'build(core): depend on the unscoped ooxml-core package', '-m',
		'@christophervr/ooxml-core is now published as ooxml-core. pptx-viewer-core@4.9.3, the one version that depends on the scoped name, is deprecated (npm will not unpublish it: the package has dependents); this releases 4.9.4.',
		'-m', $Trailer)
	Invoke-Change $Pptx 'git' @('push', 'origin', 'HEAD:main')
	Set-ReleaseWorkflow $Repos.pptx $true
	Invoke-Change $Pptx 'gh' @('workflow', 'run', 'release.yml', '-R', $Repos.pptx, '--ref', 'main')
}

$selected = if ($Only) { @($Only) } else { $Steps[$Steps.IndexOf($From)..($Steps.Count - 1)] }
if ($selected -notcontains 'preflight') { Step-Preflight }
foreach ($step in $selected) {
	Write-Host "`n=== $step ===" -ForegroundColor Magenta
	& "Step-$((Get-Culture).TextInfo.ToTitleCase($step))"
}
Write-Host "`nDone. Watch the pptx-viewer release: gh run list -R ChristopherVR/pptx-viewer -w release.yml" -ForegroundColor Green
