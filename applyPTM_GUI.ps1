Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

<#
.SYNOPSIS
    Script GUI para aplicação remota de PTMs da TOTVS
.DESCRIPTION
    Roda no micro local, permite selecionar PTMs e aplica no servidor remoto.
#>

# ====================== CONFIGURAÇÕES ======================
$config = @{
    # Servidor Protheus
    ServidorProtheus    = "127.0.0.1"
    
    # Caminhos NO SERVIDOR (remotos)
    CaminhoProtheusRemoto = "D:\TOTVS_APP\HOMOLOGATU\TOTVS_11\Protheus\bin\appserver_homologatu_compilacao"
    CaminhoProtheusExe    = "D:\TOTVS_APP\HOMOLOGATU\TOTVS_11\Protheus\bin\smartclient\smartclient.exe"  # AJUSTE AQUI COM O NOME CORRETO
    
    # AppServer
    IPAppServer           = "127.0.0.1"
    PortaAppServer        = "3909"
    NomeServico           = "TotvsProtheusHomologAtuCompilacao"
    Ambiente              = "COMPILACAO"
    
    # Backup e Log
    CaminhoBackupRemoto   = "D:\TOTVS_APP\HOMOLOGATU\TOTVS_11\Protheus\PTM\bkp_bat"
    CaminhoLogRemoto      = "D:\TOTVS_APP\HOMOLOGATU\TOTVS_11\Protheus\PTM\bkp_log"
    
    # Pasta temporária NO SERVIDOR para receber os PTMs
    PastaTempRemota       = "D:\TOTVS_APP\HOMOLOGATU\TOTVS_11\Protheus\PTM\temp"
    
    # Pasta local padrão para buscar PTMs (seu micro ou rede)
    PastaLocalDefault     = "C:\totvs\PTMs"
    
    # Credenciais (deixe $null para pedir na hora)
    Credencial            = $null
}

# ====================== FUNÇÕES AUXILIARES ======================
function Write-Status {
    param([string]$Mensagem, [string]$Tipo = "INFO")
    
    $cor = switch($Tipo) {
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Cyan" }
    }
    Write-Host "[$Tipo] $Mensagem" -ForegroundColor $cor
}

# ====================== TELA DE SELEÇÃO ======================
function Show-PTMSelectionWindow {
    param([string]$PastaDefault)

    $nullResult = [PSCustomObject]@{
        PastaPTM = $null
        Arquivos = @()
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Aplicação Remota de PTMs - TOTVS"
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = "CenterScreen"

    # ===== Servidor =====
    $lblServidor = New-Object System.Windows.Forms.Label
    $lblServidor.Text = "Servidor Protheus:"
    $lblServidor.AutoSize = $true
    $lblServidor.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($lblServidor)

    $txtServidor = New-Object System.Windows.Forms.TextBox
    $txtServidor.Location = New-Object System.Drawing.Point(10, 30)
    $txtServidor.Size = New-Object System.Drawing.Size(660, 20)
    $txtServidor.Text = "$($config.ServidorProtheus):$($config.PortaAppServer) - Ambiente: $($config.Ambiente)"
    $txtServidor.ReadOnly = $true
    $txtServidor.BackColor = [System.Drawing.Color]::LightGray
    $form.Controls.Add($txtServidor)

    # ===== Pasta PTMs =====
    $lblPasta = New-Object System.Windows.Forms.Label
    $lblPasta.Text = "Pasta dos PTMs (local ou rede):"
    $lblPasta.AutoSize = $true
    $lblPasta.Location = New-Object System.Drawing.Point(10, 60)
    $form.Controls.Add($lblPasta)

    $txtPasta = New-Object System.Windows.Forms.TextBox
    $txtPasta.Location = New-Object System.Drawing.Point(10, 80)
    $txtPasta.Size = New-Object System.Drawing.Size(600, 20)
    $txtPasta.Text = $PastaDefault
    $form.Controls.Add($txtPasta)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "..."
    $btnBrowse.Width = 50
    $btnBrowse.Location = New-Object System.Drawing.Point(620, 78)
    $form.Controls.Add($btnBrowse)

    # ===== Lista PTMs =====
    $lblLista = New-Object System.Windows.Forms.Label
    $lblLista.Text = "Selecione os PTMs que deseja aplicar:"
    $lblLista.AutoSize = $true
    $lblLista.Location = New-Object System.Drawing.Point(10, 110)
    $form.Controls.Add($lblLista)

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(10, 130)
    $checkedListBox.Size = New-Object System.Drawing.Size(660, 370)
    $checkedListBox.CheckOnClick = $true
    $form.Controls.Add($checkedListBox)

    # ===== Botões =====
    $btnAplicar = New-Object System.Windows.Forms.Button
    $btnAplicar.Text = "Aplicar no Servidor"
    $btnAplicar.Width = 140
    $btnAplicar.Location = New-Object System.Drawing.Point(400, 515)
    $btnAplicar.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnAplicar
    $form.Controls.Add($btnAplicar)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Width = 100
    $btnCancelar.Location = New-Object System.Drawing.Point(550, 515)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancelar
    $form.Controls.Add($btnCancelar)

    # ===== Atualizar lista =====
    $updateList = {
        param([string]$pasta)
        $checkedListBox.Items.Clear()
        if (-not (Test-Path $pasta)) { return }
        try {
            $files = Get-ChildItem -Path $pasta -Filter *.ptm -File -ErrorAction SilentlyContinue
            foreach ($f in $files) {
                [void]$checkedListBox.Items.Add($f.Name, $true)
            }
        } catch {}
    }

    $btnBrowse.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        if (Test-Path $txtPasta.Text) { $fbd.SelectedPath = $txtPasta.Text }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtPasta.Text = $fbd.SelectedPath
            & $updateList $fbd.SelectedPath
        }
    })

    & $updateList $txtPasta.Text

    $dialogResult = $form.ShowDialog()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        return $nullResult
    }

    $selecionados = @()
    for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
        if ($checkedListBox.GetItemChecked($i)) {
            $nome = $checkedListBox.Items[$i]
            $full = Join-Path $txtPasta.Text $nome
            if (Test-Path $full) {
                $selecionados += (Get-Item $full)
            }
        }
    }

    if ($selecionados.Count -eq 0) { return $nullResult }

    return [PSCustomObject]@{
        PastaPTM = $txtPasta.Text
        Arquivos = $selecionados
    }
}

# ====================== APLICAÇÃO REMOTA ======================
function Invoke-RemotePTMApplication {
    param(
        [System.IO.FileInfo[]]$PTMs,
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Status "Conectando ao servidor $($config.ServidorProtheus)..."

    try {
        $session = New-PSSession -ComputerName $config.ServidorProtheus -Credential $Credential -ErrorAction Stop
        Write-Status "Sessão remota estabelecida com sucesso" "SUCCESS"

        Write-Status "Criando pasta temporária no servidor..."
        Invoke-Command -Session $session -ScriptBlock {
            param($pasta)
            if (-not (Test-Path $pasta)) {
                New-Item -ItemType Directory -Path $pasta -Force | Out-Null
            }
        } -ArgumentList $config.PastaTempRemota

        Write-Status "Copiando $($PTMs.Count) PTM(s) para o servidor..."
        foreach ($ptm in $PTMs) {
            $nomeArquivo = $ptm.Name
            $caminhoRemoto = Invoke-Command -Session $session -ScriptBlock {
                param($pastaTemp, $nome)
                Join-Path $pastaTemp $nome
            } -ArgumentList $config.PastaTempRemota, $nomeArquivo
            
            Copy-Item -Path $ptm.FullName -Destination $caminhoRemoto -ToSession $session -Force
            Write-Status "  → $($ptm.Name) copiado" "SUCCESS"
        }

        Write-Status "Executando aplicação de PTMs no servidor..."
        
        $resultado = Invoke-Command -Session $session -ScriptBlock {
            param($cfg, $ptmNames)

            $log = @()
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $backupPath = Join-Path $cfg.CaminhoBackupRemoto "PTMs_$timestamp"
            $logPath = Join-Path $cfg.CaminhoLogRemoto "AplicacaoPTM_$timestamp.log"

            function Add-Log {
                param([string]$msg, [string]$tipo = "INFO")
                $script:log += "[$tipo] $msg"
                $logLine = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$tipo] $msg"
                
                $logDir = Split-Path $logPath -Parent
                if (-not (Test-Path $logDir)) {
                    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
                }
                Add-Content -Path $logPath -Value $logLine
            }

            Add-Log "=========================================="
            Add-Log "Iniciando aplicação de PTMs"
            Add-Log "Ambiente: $($cfg.Ambiente)"
            Add-Log "AppServer: $($cfg.IPAppServer):$($cfg.PortaAppServer)"
            Add-Log "Total de PTMs: $($ptmNames.Count)"
            Add-Log "=========================================="

            Add-Log "Parando serviço $($cfg.NomeServico)..."
            try {
                $svc = Get-Service -Name $cfg.NomeServico -ErrorAction Stop
                if ($svc.Status -eq "Running") {
                    Stop-Service -Name $cfg.NomeServico -Force
                    Start-Sleep -Seconds 5
                    
                    $svc = Get-Service -Name $cfg.NomeServico
                    if ($svc.Status -eq "Stopped") {
                        Add-Log "Serviço parado com sucesso" "SUCCESS"
                    } else {
                        Add-Log "Serviço não parou completamente (Status: $($svc.Status))" "WARNING"
                    }
                } else {
                    Add-Log "Serviço já estava parado" "WARNING"
                }
            } catch {
                Add-Log "Erro ao parar serviço: $_" "ERROR"
                return @{ Sucesso = $false; Log = $script:log }
            }

            Add-Log "Criando backup do RPO em $backupPath..."
            try {
                New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
                
                $rpoPath = Join-Path $cfg.CaminhoProtheusRemoto "*.rpo"
                $rpoFiles = Get-ChildItem -Path $rpoPath -ErrorAction SilentlyContinue
                
                foreach ($rpo in $rpoFiles) {
                    Copy-Item -Path $rpo.FullName -Destination (Join-Path $backupPath $rpo.Name) -Force
                    Add-Log "Backup do RPO: $($rpo.Name)" "SUCCESS"
                }
                
                Add-Log "Backup concluído" "SUCCESS"
            } catch {
                Add-Log "Erro no backup: $_" "ERROR"
                Start-Service -Name $cfg.NomeServico -ErrorAction SilentlyContinue
                return @{ Sucesso = $false; Log = $script:log }
            }

            Add-Log "Aplicando PTMs individualmente usando Protheus.exe..."
            
            $ptmCounter = 0
            $ptmTotal = $ptmNames.Count
            
            foreach ($ptmName in $ptmNames) {
                $ptmCounter++
                
                Add-Log "=========================================="
                Add-Log "Aplicando PTM $ptmCounter de $ptmTotal : $ptmName"
                Add-Log "=========================================="
                
                $ptmPath = Join-Path $cfg.PastaTempRemota $ptmName
                
                if (-not (Test-Path $ptmPath)) {
                    Add-Log "PTM não encontrado: $ptmName" "WARNING"
                    continue
                }
                
                $protheusExe = $cfg.CaminhoProtheusExe
                $arguments = @(
                    "-console",
                    "-appserver=$($cfg.IPAppServer)",
                    "-port=$($cfg.PortaAppServer)",
                    "-compile",
                    "-applypatch",
                    "-env=$($cfg.Ambiente)",
                    "-files=$ptmPath"
                )
                
                Add-Log "Executando: $protheusExe $($arguments -join ' ')"
                
                try {
                    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $processInfo.FileName = $protheusExe
                    $processInfo.Arguments = $arguments -join " "
                    $processInfo.RedirectStandardOutput = $true
                    $processInfo.RedirectStandardError = $true
                    $processInfo.UseShellExecute = $false
                    $processInfo.CreateNoWindow = $true
                    $processInfo.WorkingDirectory = Split-Path $protheusExe -Parent
                    
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $processInfo
                    
                    $outputBuilder = New-Object System.Text.StringBuilder
                    $errorBuilder = New-Object System.Text.StringBuilder
                    
                    $outputEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action {
                        if (-not [string]::IsNullOrEmpty($EventArgs.Data)) {
                            $Event.MessageData.AppendLine($EventArgs.Data)
                        }
                    } -MessageData $outputBuilder
                    
                    $errorEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action {
                        if (-not [string]::IsNullOrEmpty($EventArgs.Data)) {
                            $Event.MessageData.AppendLine($EventArgs.Data)
                        }
                    } -MessageData $errorBuilder
                    
                    [void]$process.Start()
                    $process.BeginOutputReadLine()
                    $process.BeginErrorReadLine()
                    
                    $timeout = 300000 
                    if (-not $process.WaitForExit($timeout)) {
                        Add-Log "Timeout ao aplicar PTM $ptmName (5 minutos)" "ERROR"
                        $process.Kill()
                        Unregister-Event -SourceIdentifier $outputEvent.Name
                        Unregister-Event -SourceIdentifier $errorEvent.Name
                        continue
                    }
                    
                    Unregister-Event -SourceIdentifier $outputEvent.Name
                    Unregister-Event -SourceIdentifier $errorEvent.Name
                    
                    $output = $outputBuilder.ToString()
                    $errorOutput = $errorBuilder.ToString()
                    
                    Add-Log "Exit Code: $($process.ExitCode)"
                    
                    if ($output) {
                        Add-Log "=== Saída do Protheus.exe para $ptmName ==="
                        foreach ($line in $output -split "`n") {
                            if ($line.Trim()) { Add-Log $line }
                        }
                    }
                    
                    if ($errorOutput) {
                        Add-Log "=== Erros do Protheus.exe para $ptmName ===" "WARNING"
                        foreach ($line in $errorOutput -split "`n") {
                            if ($line.Trim()) { Add-Log $line "WARNING" }
                        }
                    }
                    
                    if ($process.ExitCode -eq 0) {
                        Add-Log "PTM $ptmName aplicado com sucesso!" "SUCCESS"
                    } else {
                        Add-Log "Protheus.exe retornou código de erro: $($process.ExitCode) para $ptmName" "WARNING"
                    }
                    
                } catch {
                    Add-Log "Erro ao executar Protheus.exe para $ptmName : $_" "ERROR"
                }
                
                Add-Log "Aguardando 3 segundos antes do próximo PTM..."
                Start-Sleep -Seconds 3
            }

            Add-Log "=========================================="
            Add-Log "Todos os PTMs foram processados"
            Add-Log "=========================================="

            Add-Log "Iniciando serviço $($cfg.NomeServico)..."
            try {
                Start-Service -Name $cfg.NomeServico
                Start-Sleep -Seconds 10
                
                $svc = Get-Service -Name $cfg.NomeServico
                if ($svc.Status -eq "Running") {
                    Add-Log "Serviço iniciado com sucesso" "SUCCESS"
                } else {
                    Add-Log "Serviço não iniciou corretamente (Status: $($svc.Status))" "WARNING"
                }
            } catch {
                Add-Log "Erro ao iniciar serviço: $_" "ERROR"
            }
			
            Add-Log "Parando serviço $($cfg.NomeServico)..."
            try {
                $svc = Get-Service -Name $cfg.NomeServico -ErrorAction Stop
                if ($svc.Status -eq "Running") {
                    Stop-Service -Name $cfg.NomeServico -Force
                    Start-Sleep -Seconds 5
                    
                    $svc = Get-Service -Name $cfg.NomeServico
                    if ($svc.Status -eq "Stopped") {
                        Add-Log "Serviço parado com sucesso" "SUCCESS"
                    } else {
                        Add-Log "Serviço não parou completamente (Status: $($svc.Status))" "WARNING"
                    }
                } else {
                    Add-Log "Serviço já estava parado" "WARNING"
                }
            } catch {
                Add-Log "Erro ao parar serviço: $_" "ERROR"
                return @{ Sucesso = $false; Log = $script:log }
            }

            Add-Log "=========================================="
            Add-Log "Processo finalizado"
            Add-Log "Log salvo em: $logPath"
            Add-Log "Backup salvo em: $backupPath"
            Add-Log "=========================================="

            return @{
                Sucesso = $true
                Log = $script:log
                CaminhoLog = $logPath
                CaminhoBackup = $backupPath
            }

        } -ArgumentList $config, ($PTMs | ForEach-Object { $_.Name })

        Write-Host "`n========== LOG DO SERVIDOR ==========" -ForegroundColor Cyan
        foreach ($linha in $resultado.Log) {
            Write-Host $linha
        }
        Write-Host "====================================`n" -ForegroundColor Cyan

        if ($resultado.Sucesso) {
            Write-Status "Aplicação concluída com sucesso!" "SUCCESS"
            Write-Status "Log: $($resultado.CaminhoLog)" "SUCCESS"
            Write-Status "Backup: $($resultado.CaminhoBackup)" "SUCCESS"
        }

        Write-Status "Limpando pasta temporária..."
        Invoke-Command -Session $session -ScriptBlock {
            param($pasta)
            Remove-Item -Path "$pasta\*.ptm" -Force -ErrorAction SilentlyContinue
        } -ArgumentList $config.PastaTempRemota

        Remove-PSSession -Session $session

    } catch {
        Write-Status "Erro na execução remota: $($_.Exception.Message)" "ERROR"
        return $false
    }

    return $true
}

# ====================== MAIN ======================
function Main {
    Write-Host "========================================"   -ForegroundColor Cyan
    Write-Host "  Aplicação Remota de PTMs - TOTVS"         -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    if (-not $config.Credencial) {
        Write-Status "Informe as credenciais de administrador do servidor $($config.ServidorProtheus)"
        $config.Credencial = Get-Credential -Message "Credenciais do servidor $($config.ServidorProtheus)"
        if (-not $config.Credencial) {
            Write-Status "Credenciais não fornecidas. Abortando." "ERROR"
            return
        }
    }

    $result = Show-PTMSelectionWindow -PastaDefault $config.PastaLocalDefault

    if (-not $result.Arquivos -or $result.Arquivos.Count -eq 0) {
        Write-Status "Nenhum PTM selecionado. Operação cancelada." "WARNING"
        return
    }

    Write-Status "PTMs selecionados: $($result.Arquivos.Count)"
    foreach ($ptm in $result.Arquivos) {
        Write-Status "  → $($ptm.Name)"
    }

    $confirmacao = [System.Windows.Forms.MessageBox]::Show(
        "Deseja aplicar $($result.Arquivos.Count) PTM(s) no servidor $($config.ServidorProtheus)?`n`nAmbiente: $($config.Ambiente)`nPorta: $($config.PortaAppServer)`n`nOs PTMs serão aplicados UM POR VEZ.`n`nEsta operação irá parar o serviço TOTVS temporariamente.",
        "Confirmação",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmacao -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Status "Operação cancelada pelo usuário." "WARNING"
        return
    }

    Invoke-RemotePTMApplication -PTMs $result.Arquivos -Credential $config.Credencial

    Write-Host "`nPressione qualquer tecla para sair..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ====================== EXECUÇÃO ======================
Main