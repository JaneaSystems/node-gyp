const log = require('npmlog')
  , execSync = require('child_process').execSync
  , path = require('path')

var hasCache = false
  , cache = null

function findVS2017() {
  if (hasCache)
    return cache

  hasCache = true

  const ps = 'PowerShell -ExecutionPolicy Unrestricted -Command '
  const psClsidTest = ps + 'Test-Path \'Registry::HKCR\\CLSID\\{177F0C4A-1CD3-4DE7-A32C-71DBBB9FA36D}\' 2>&1'

  try {
    const hasClsidRaw = execSync(psClsidTest, { encoding: 'utf8' })
    log.silly('find vs2017', 'hasClsidRaw:', hasClsidRaw)
    const hasClsid = hasClsidRaw.trim() === 'True'
    if (!hasClsid)
      log.verbose('find vs2017', 'could not find VS2017 installer class in registry')
  } catch (e) {
    log.verbose('find vs2017', e)
  }

  const csFile = path.join(__dirname, 'Find-VS2017.cs')
  const psQuery = ps + '"&{Add-Type -Path \'' + csFile + '\'; [VisualStudioConfiguration.Main]::Query()}" 2>&1'

  try {
    const vsSetupRaw = execSync(psQuery, { encoding: 'utf8' })
    log.silly('find vs2017', 'vsSetupRaw:', vsSetupRaw)
  } catch (e) {
    log.verbose('find vs2017', e)
    return cache
  }

  try {
    const vsSetup = JSON.parse(vsSetupRaw)
    log.silly('find vs2017', 'vsSetup:', vsSetup)
  } catch (e) {
    log.verbose('find vs2017', e)
    return cache
  }

  log.verbose('find vs2017', vsSetup.log)

  cache = {
    "path": vsSetup.path,
    "sdk": vsSetup.sdk
  }

  return cache
}

module.exports = findVS2017
