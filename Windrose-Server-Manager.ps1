Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinHelper {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    public static void HideProcessWindows(int pid) {
        EnumWindows((hWnd, lParam) => {
            int windowPid = 0;
            GetWindowThreadProcessId(hWnd, out windowPid);
            if (windowPid == pid && IsWindowVisible(hWnd)) ShowWindow(hWnd, 0);
            return true;
        }, IntPtr.Zero);
    }

    // --- Console input via Win32 (bypasses broken stdin pipe) ---
    [DllImport("kernel32.dll", SetLastError = true)] public static extern bool FreeConsole();
    [DllImport("kernel32.dll", SetLastError = true)] public static extern bool AttachConsole(int pid);
    [DllImport("kernel32.dll", SetLastError = true)] public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError = true)] public static extern bool AllocConsole();
    [DllImport("kernel32.dll", SetLastError = true)] public static extern IntPtr GetConsoleWindow();

    [StructLayout(LayoutKind.Sequential)]
    public struct KEY_EVENT_RECORD {
        public int bKeyDown;
        public short wRepeatCount;
        public short wVirtualKeyCode;
        public short wVirtualScanCode;
        public char UnicodeChar;
        public int dwControlKeyState;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct INPUT_RECORD {
        [FieldOffset(0)] public short EventType;
        [FieldOffset(4)] public KEY_EVENT_RECORD KeyEvent;
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool WriteConsoleInput(
        IntPtr hConsoleInput, INPUT_RECORD[] lpBuffer, int nLength, out int written);

    public static bool SendConsoleCommand(int pid, string command) {
        FreeConsole();
        if (!AttachConsole(pid)) return false;
        try {
            IntPtr hInput = GetStdHandle(-10); // STD_INPUT_HANDLE
            if (hInput == IntPtr.Zero || hInput == new IntPtr(-1)) return false;

            string full = command + "\r";
            var records = new INPUT_RECORD[full.Length * 2];
            int idx = 0;
            foreach (char c in full) {
                short vk = (c == '\r') ? (short)0x0D : (short)0;
                records[idx++] = new INPUT_RECORD {
                    EventType = 1,
                    KeyEvent = new KEY_EVENT_RECORD {
                        bKeyDown = 1, wRepeatCount = 1,
                        wVirtualKeyCode = vk, UnicodeChar = c
                    }
                };
                records[idx++] = new INPUT_RECORD {
                    EventType = 1,
                    KeyEvent = new KEY_EVENT_RECORD {
                        bKeyDown = 0, wRepeatCount = 1,
                        wVirtualKeyCode = vk, UnicodeChar = c
                    }
                };
            }
            int written;
            return WriteConsoleInput(hInput, records, records.Length, out written);
        } finally {
            FreeConsole();
            // Restore a hidden console for PowerShell so it doesn't crash
            AllocConsole();
            IntPtr cw = GetConsoleWindow();
            if (cw != IntPtr.Zero) ShowWindow(cw, 0);
        }
    }
}
"@

$AppVersion  = "1.21"
$UpdateUrl   = "https://raw.githubusercontent.com/psbrowand/Windrose-Server-Manager/main/Windrose-Server-Manager.ps1"

$PatchNotes = [ordered]@{
    "1.21" = @(
        "Moved Player History from Tools tab to Dashboard -- now side-by-side with Connected Players",
        "Dashboard split into two-column layout: Connected Players (left) and Player History (right)"
    )
    "1.20" = @(
        "Fixed console command crash -- restored PowerShell console after Win32 detach/attach",
        "Fixed backup status text not updating after successful backup",
        "Added auto-backup feature with selectable interval (1h, 4h, 8h, 16h, 24h)",
        "Auto-backup shows next scheduled backup time in the Tools tab"
    )
    "1.19" = @(
        "Fixed console commands -- replaced broken stdin pipe with Win32 WriteConsoleInput",
        "Updated command syntax to match Windrose server (save world, list players, etc.)",
        "Fixed player disconnect detection -- now catches SaidFarewell and DisconnectAccount events",
        "Players no longer appear stuck online after ungraceful disconnects (crash, timeout)",
        "Console tab now shows server log output so command responses are visible",
        "Added Save on Stop checkbox -- saves world before stopping the server",
        "Server Info button renamed to Show Logs (uses the 'logs' command)"
    )
    "1.18" = @(
        "Performance: Log viewer now appends new lines instead of rebuilding every 3 seconds",
        "Performance: Player list only redraws when players join or leave (no more flicker)",
        "Performance: Cached shared brushes and fonts -- eliminates thousands of object allocations per minute",
        "Refactored log filter buttons into a single function"
    )
    "1.17" = @(
        "Fixed ComboBox header (selected value area) background -- was white, now dark",
        "Full ControlTemplate applied to dropdowns for consistent dark theming throughout"
    )
    "1.16" = @(
        "Fixed ComboBox dropdown popup background using DropDownOpened event",
        "Popup border now correctly shows dark background when dropdown opens"
    )
    "1.15" = @(
        "Fixed ComboBox dropdown popup background (was white, now dark)",
        "Added Reload Saved Config button to Config tab -- discards unsaved changes"
    )
    "1.14" = @(
        "Added Patch Notes button in Tools tab",
        "Shows a scrollable history of changes for each version"
    )
    "1.13" = @(
        "Fixed Update Now button hanging indefinitely on download",
        "Download job now uses script-scoped variables so the timer can track it",
        "Check for Updates now cancels and resets any stuck download"
    )
    "1.12" = @(
        "Fixed launch crash caused by special characters (encoding issue with PowerShell 5.1)",
        "App version number now correctly tracked and pushed to GitHub"
    )
    "1.11" = @(
        "Install tab replaced with a 5-step setup wizard",
        "Step 1: auto-checks whether Windrose is installed on Steam",
        "Step 2: install server files (existing flow, reorganized)",
        "Step 3: name your server and set max players before first launch",
        "Step 4: port forwarding reference card with copy-to-clipboard",
        "Step 5: Go to Dashboard button"
    )
    "1.10" = @(
        "Switched to 1.XX version numbering scheme",
        "Fixed update version comparison logic",
        "Fixed max players clamp -- values above 10 now load correctly",
        "Default max players changed from 4 to 10"
    )
    "1.09" = @(
        "Added manual Refresh button to Dashboard player list",
        "Player list auto-resyncs from full log every 30 seconds to fix stale entries"
    )
    "1.08" = @(
        "Config tab now pre-populates world settings on startup",
        "World config save now writes correct nested JSON structure",
        "Console commands re-enabled -- server launched with -log so UE5 reads stdin",
        "Server console window hidden via Win32 API after startup"
    )
    "1.07" = @(
        "Added self-update feature (Check for Updates / Update Now in Tools tab)",
        "Added right-click kick and ban on the Dashboard player list",
        "Dashboard player count fixed (was always showing 0)",
        "Max players slider extended from 10 to 20"
    )
}

$ServerDir      = $PSScriptRoot
$ServerExe      = "$ServerDir\WindroseServer.exe"
$ServerExeDirect= "$ServerDir\R5\Binaries\Win64\WindroseServer-Win64-Shipping.exe"
$ConfigPath  = "$ServerDir\R5\ServerDescription.json"
$LogPath     = "$ServerDir\R5\Saved\Logs\R5.log"
$SavesBase   = "$ServerDir\R5\Saved\SaveProfiles"
$BackupDir   = "$ServerDir\Backups"
$HistoryFile = "$ServerDir\player_history.txt"
if (-not (Test-Path $BackupDir)) { New-Item $BackupDir -ItemType Directory -Force | Out-Null }

[xml]$Xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Windrose Server Manager"
    Height="780" Width="700"
    MinHeight="640" MinWidth="560"
    Background="#0F1923"
    WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style x:Key="BaseBtn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="14,6"/>
      <Setter Property="Margin" Value="3"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="4"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="SmallBtn" TargetType="Button" BasedOn="{StaticResource BaseBtn}">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="8,4"/>
    </Style>
    <Style x:Key="DarkInput" TargetType="TextBox">
      <Setter Property="Background" Value="#1A2736"/>
      <Setter Property="Foreground" Value="#D0D8E4"/>
      <Setter Property="BorderBrush" Value="#2A3E55"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="CaretBrush" Value="White"/>
    </Style>
    <Style x:Key="DarkCheck" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#B0BEC5"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Margin" Value="0,4,0,4"/>
    </Style>
    <Style x:Key="DarkCombo" TargetType="ComboBox">
      <Setter Property="Background" Value="#1A2736"/>
      <Setter Property="Foreground" Value="#D0D8E4"/>
      <Setter Property="BorderBrush" Value="#2A3E55"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="ItemContainerStyle">
        <Setter.Value>
          <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#1A2736"/>
            <Setter Property="Foreground" Value="#D0D8E4"/>
            <Setter Property="Padding" Value="8,5"/>
            <Style.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#1E3348"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="#1E3348"/>
                <Setter Property="Foreground" Value="#D4A843"/>
              </Trigger>
            </Style.Triggers>
          </Style>
        </Setter.Value>
      </Setter>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid x:Name="templateRoot">
              <ToggleButton x:Name="toggleButton"
                Background="{TemplateBinding Background}"
                BorderBrush="{TemplateBinding BorderBrush}"
                BorderThickness="{TemplateBinding BorderThickness}"
                IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                Focusable="false" ClickMode="Press">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition/>
                          <ColumnDefinition Width="20"/>
                        </Grid.ColumnDefinitions>
                        <Border Grid.Column="1" BorderBrush="#2A3E55" BorderThickness="1,0,0,0">
                          <Path Fill="#D0D8E4" HorizontalAlignment="Center" VerticalAlignment="Center"
                                Data="M 0 0 L 4 4 L 8 0 Z"/>
                        </Border>
                      </Grid>
                    </Border>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter
                Content="{TemplateBinding SelectionBoxItem}"
                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                ContentStringFormat="{TemplateBinding SelectionBoxItemStringFormat}"
                HorizontalAlignment="Left"
                VerticalAlignment="Center"
                Margin="{TemplateBinding Padding}"
                IsHitTestVisible="false"
                TextBlock.Foreground="{TemplateBinding Foreground}"/>
              <Popup x:Name="PART_Popup"
                IsOpen="{TemplateBinding IsDropDownOpen}"
                Placement="Bottom"
                AllowsTransparency="true"
                Focusable="false"
                MinWidth="{Binding ActualWidth, ElementName=toggleButton}">
                <Border Background="#1A2736" BorderBrush="#2A3E55" BorderThickness="1"
                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                  <ScrollViewer>
                    <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"
                                    SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="SectionHead" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#D4A843"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Margin" Value="0,12,0,6"/>
    </Style>
    <Style x:Key="FieldLabel" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#8DA4B5"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="#0F1923"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Background" Value="#162330"/>
      <Setter Property="Foreground" Value="#8DA4B5"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Grid>
              <Border x:Name="TabBorder" Background="{TemplateBinding Background}"
                      Padding="{TemplateBinding Padding}">
                <ContentPresenter ContentSource="Header"
                                  HorizontalAlignment="Center"
                                  VerticalAlignment="Center"
                                  TextBlock.Foreground="{TemplateBinding Foreground}"/>
              </Border>
              <Border x:Name="ActiveLine" Height="3" VerticalAlignment="Bottom"
                      Background="#D4A843" Visibility="Collapsed"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="TabBorder" Property="Background" Value="#1E3348"/>
                <Setter Property="Foreground" Value="#D4A843"/>
                <Setter TargetName="ActiveLine" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="TabBorder" Property="Background" Value="#1A2D40"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ListBox">
      <Setter Property="Background" Value="#111E2A"/>
      <Setter Property="BorderBrush" Value="#1E3348"/>
      <Setter Property="Foreground" Value="#C0CDD8"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="2"/>
    </Style>
    <Style TargetType="Slider">
      <Setter Property="Foreground" Value="#D4A843"/>
    </Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <!-- HEADER -->
    <Border Grid.Row="0" Background="#0A1520" Padding="12,10">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Vertical">
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <Ellipse x:Name="DotStatus" Width="12" Height="12" Fill="#555" Margin="0,0,8,0" VerticalAlignment="Center"/>
            <TextBlock x:Name="TxtServerName" Text="Windrose Server" FontSize="18" FontWeight="Bold"
                       Foreground="#D4A843" VerticalAlignment="Center"/>
            <TextBlock x:Name="TxtStatus" Text="  Stopped" FontSize="13"
                       Foreground="#8DA4B5" VerticalAlignment="Center" Margin="10,0,0,0"/>
            <TextBlock x:Name="TxtUptime" Text="" FontSize="12"
                       Foreground="#607080" VerticalAlignment="Center" Margin="16,0,0,0"/>
          </StackPanel>
          <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
            <TextBlock Text="Code:" Foreground="#607080" FontSize="11" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBlock x:Name="TxtInvite" Text="--" FontSize="11" Foreground="#A0C4E0"
                       VerticalAlignment="Center" Cursor="Hand"
                       ToolTip="Click to copy invite code"/>
          </StackPanel>
        </StackPanel>
        <Button x:Name="BtnShare" Grid.Column="1" Content="Share" VerticalAlignment="Center"
                Background="#1A4A7A" Style="{StaticResource SmallBtn}"/>
      </Grid>
    </Border>
    <!-- TABS -->
    <TabControl x:Name="MainTabs" Grid.Row="1" Margin="0">
      <!-- TAB 1: DASHBOARD -->
      <TabItem Header="Dashboard">
        <Grid Margin="12">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <!-- Stats bar -->
          <Border Grid.Row="0" Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1"
                  CornerRadius="6" Padding="12,10" Margin="0,0,0,10">
            <UniformGrid Rows="1" Columns="4">
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="CPU" Foreground="#607080" FontSize="11" HorizontalAlignment="Center"/>
                <TextBlock x:Name="TxtCpu" Text="--" Foreground="#D4A843" FontSize="22"
                           FontWeight="Bold" HorizontalAlignment="Center"/>
              </StackPanel>
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="RAM" Foreground="#607080" FontSize="11" HorizontalAlignment="Center"/>
                <TextBlock x:Name="TxtRam" Text="--" Foreground="#5BA4CF" FontSize="22"
                           FontWeight="Bold" HorizontalAlignment="Center"/>
              </StackPanel>
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="PLAYERS" Foreground="#607080" FontSize="11" HorizontalAlignment="Center"/>
                <TextBlock x:Name="TxtPlayers" Text="--" Foreground="#70C48A" FontSize="22"
                           FontWeight="Bold" HorizontalAlignment="Center"/>
              </StackPanel>
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="UPTIME" Foreground="#607080" FontSize="11" HorizontalAlignment="Center"/>
                <TextBlock x:Name="TxtUptimeBig" Text="--" Foreground="#A0C4E0" FontSize="22"
                           FontWeight="Bold" HorizontalAlignment="Center"/>
              </StackPanel>
            </UniformGrid>
          </Border>
          <!-- Two-column: Connected Players | Player History -->
          <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <!-- Left: Connected Players -->
            <DockPanel Grid.Column="0" Margin="0,0,6,0">
              <Grid DockPanel.Dock="Top" Margin="0,0,0,4">
                <TextBlock Text="Connected Players" Style="{StaticResource SectionHead}" VerticalAlignment="Center"/>
                <Button x:Name="BtnRefreshPlayers" Content="Refresh" HorizontalAlignment="Right"
                        Background="#2A3E55" Style="{StaticResource SmallBtn}" VerticalAlignment="Center"/>
              </Grid>
              <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="4">
                <ListBox x:Name="PlayerList" FontSize="13" BorderThickness="0" Background="Transparent">
                  <ListBox.ItemContainerStyle>
                    <Style TargetType="ListBoxItem">
                      <Setter Property="Foreground" Value="#C0CDD8"/>
                      <Setter Property="Padding" Value="6,3"/>
                    </Style>
                  </ListBox.ItemContainerStyle>
                </ListBox>
              </Border>
            </DockPanel>
            <!-- Right: Player History -->
            <DockPanel Grid.Column="1" Margin="6,0,0,0">
              <Grid DockPanel.Dock="Top" Margin="0,0,0,4">
                <TextBlock Text="Player History" Style="{StaticResource SectionHead}" VerticalAlignment="Center"/>
                <Button x:Name="BtnClearHistory" Content="Clear" HorizontalAlignment="Right"
                        Background="#5A2020" Style="{StaticResource SmallBtn}" VerticalAlignment="Center"/>
              </Grid>
              <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="4">
                <ListBox x:Name="HistoryList" FontSize="11" FontFamily="Consolas" BorderThickness="0" Background="Transparent">
                  <ListBox.ItemContainerStyle>
                    <Style TargetType="ListBoxItem">
                      <Setter Property="Foreground" Value="#C0CDD8"/>
                      <Setter Property="Padding" Value="4,2"/>
                    </Style>
                  </ListBox.ItemContainerStyle>
                </ListBox>
              </Border>
            </DockPanel>
          </Grid>
          <!-- Checkboxes -->
          <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
            <CheckBox x:Name="ChkAutoRestart" Content="Auto-restart if crashed"
                      Style="{StaticResource DarkCheck}" Margin="0,0,20,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="ChkSaveOnStop" Content="Save world on stop"
                      Style="{StaticResource DarkCheck}" IsChecked="True" VerticalAlignment="Center"/>
          </StackPanel>
        </Grid>
      </TabItem>
      <!-- TAB 2: CONFIG -->
      <TabItem Header="Config">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#0F1923">
          <StackPanel Margin="14,10">
            <TextBlock Text="Server Settings" Style="{StaticResource SectionHead}" Margin="0,0,0,6"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="Server Name" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="CfgName" Grid.Column="1" Style="{StaticResource DarkInput}"/>
                </Grid>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="40"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="Max Players" Style="{StaticResource FieldLabel}"/>
                  <Slider x:Name="CfgMaxPlayers" Grid.Column="1" Minimum="1" Maximum="20"
                          TickFrequency="1" IsSnapToTickEnabled="True" Value="10"
                          VerticalAlignment="Center"/>
                  <TextBlock x:Name="TxtMaxPlayersVal" Grid.Column="2" Text="10"
                             Foreground="#D4A843" FontWeight="Bold" VerticalAlignment="Center"
                             HorizontalAlignment="Center"/>
                </Grid>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="Password" Style="{StaticResource FieldLabel}"/>
                  <CheckBox x:Name="CfgPasswordEnabled" Grid.Column="1" Content="Enable"
                            Style="{StaticResource DarkCheck}" Margin="0,0,10,0"/>
                  <TextBox x:Name="CfgPassword" Grid.Column="2" Style="{StaticResource DarkInput}"
                           IsEnabled="False"/>
                </Grid>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="Proxy Address" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="CfgProxy" Grid.Column="1" Style="{StaticResource DarkInput}"/>
                </Grid>
              </StackPanel>
            </Border>
            <TextBlock Text="World Settings" Style="{StaticResource SectionHead}" Margin="0,12,0,6"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="Difficulty Preset" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CfgPreset" Grid.Column="1" Style="{StaticResource DarkCombo}">
                    <ComboBoxItem Content="Easy"/>
                    <ComboBoxItem Content="Medium"/>
                    <ComboBoxItem Content="Hard"/>
                    <ComboBoxItem Content="Custom"/>
                  </ComboBox>
                </Grid>
                <StackPanel x:Name="PanelCustom" Visibility="Collapsed" Margin="0,0,0,8">
                  <Border Background="#0D1820" CornerRadius="4" Padding="10" Margin="0,0,0,4">
                    <StackPanel>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Mob Health" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlMobHealth" Grid.Column="1" Minimum="0.2" Maximum="5.0"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValMobHealth" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Mob Damage" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlMobDamage" Grid.Column="1" Minimum="0.2" Maximum="5.0"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValMobDamage" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Ship Health" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlShipHealth" Grid.Column="1" Minimum="0.4" Maximum="5.0"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValShipHealth" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Ship Damage" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlShipDamage" Grid.Column="1" Minimum="0.2" Maximum="2.5"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValShipDamage" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Boarding" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlBoarding" Grid.Column="1" Minimum="0.2" Maximum="5.0"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValBoarding" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Coop Stats" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlCoopStats" Grid.Column="1" Minimum="0.0" Maximum="2.0"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValCoopStats" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                      <Grid Margin="0,2">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="130"/>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Coop Ship" Style="{StaticResource FieldLabel}"/>
                        <Slider x:Name="SlCoopShip" Grid.Column="1" Minimum="0.0" Maximum="2.0"
                                TickFrequency="0.1" Value="1.0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="ValCoopShip" Grid.Column="2" Text="1.0"
                                   Foreground="#D4A843" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="11"/>
                      </Grid>
                    </StackPanel>
                  </Border>
                </StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="Combat Diff." Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CfgCombatDiff" Grid.Column="1" Style="{StaticResource DarkCombo}">
                    <ComboBoxItem Content="Easy"/>
                    <ComboBoxItem Content="Normal"/>
                    <ComboBoxItem Content="Hard"/>
                  </ComboBox>
                </Grid>
                <CheckBox x:Name="CfgCoopQuests" Content="Coop Quests" Style="{StaticResource DarkCheck}"/>
                <CheckBox x:Name="CfgEasyExplore" Content="Easy Exploration" Style="{StaticResource DarkCheck}"/>
              </StackPanel>
            </Border>
            <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
              <Button x:Name="BtnSaveConfig" Content="Save Config" Background="#1A6B3A" Style="{StaticResource BaseBtn}"/>
              <Button x:Name="BtnReloadConfig" Content="Reload Saved Config" Background="#2A3E55" Style="{StaticResource BaseBtn}"/>
              <Button x:Name="BtnOpenWorldJson" Content="Open World JSON" Background="#1A3A6B" Style="{StaticResource BaseBtn}"/>
            </StackPanel>
            <TextBlock x:Name="TxtConfigStatus" Text="" Foreground="#70C48A" FontSize="11" Margin="4,6,0,0"/>
          </StackPanel>
        </ScrollViewer>
      </TabItem>
      <!-- TAB 3: LOG -->
      <TabItem Header="Log">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,6">
            <Button x:Name="BtnFAll" Content="All" Background="#1A4A7A" Style="{StaticResource SmallBtn}"/>
            <Button x:Name="BtnFPlayers" Content="Players" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
            <Button x:Name="BtnFWarn" Content="Warnings" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
            <Button x:Name="BtnFErrors" Content="Errors" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
            <CheckBox x:Name="ChkAutoScroll" Content="Auto-scroll" Style="{StaticResource DarkCheck}"
                      IsChecked="True" Margin="16,0,0,0" VerticalAlignment="Center"/>
          </StackPanel>
          <ListBox x:Name="LogViewer" Grid.Row="1" FontFamily="Consolas" FontSize="11"
                   VirtualizingStackPanel.IsVirtualizing="True"
                   VirtualizingStackPanel.VirtualizationMode="Recycling"
                   ScrollViewer.HorizontalScrollBarVisibility="Auto">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Padding" Value="2,1"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
              </Style>
            </ListBox.ItemContainerStyle>
          </ListBox>
        </Grid>
      </TabItem>
      <!-- TAB 4: CONSOLE -->
      <TabItem Header="Console">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <ListBox x:Name="ConsoleOutput" Grid.Row="0" FontFamily="Consolas" FontSize="11"
                   VirtualizingStackPanel.IsVirtualizing="True"
                   VirtualizingStackPanel.VirtualizationMode="Recycling"
                   ScrollViewer.HorizontalScrollBarVisibility="Auto"
                   Margin="0,0,0,6">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Padding" Value="2,1"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
              </Style>
            </ListBox.ItemContainerStyle>
          </ListBox>
          <WrapPanel Grid.Row="1" Margin="0,0,0,6">
            <Button x:Name="BtnCmdSave" Content="Save World" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
            <Button x:Name="BtnCmdPlayers" Content="List Players" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
            <Button x:Name="BtnCmdInfo" Content="Show Logs" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
            <Button x:Name="BtnCmdQuit" Content="Quit Server" Background="#5A2020" Style="{StaticResource SmallBtn}"/>
          </WrapPanel>
          <Grid Grid.Row="2" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtCommand" Grid.Column="0" Style="{StaticResource DarkInput}"
                     Tag="Enter command..." Margin="0,0,4,0"/>
            <Button x:Name="BtnSendCmd" Grid.Column="1" Content="Send"
                    Background="#1A4A7A" Style="{StaticResource BaseBtn}"/>
          </Grid>
          <TextBlock x:Name="TxtConsoleStatus" Grid.Row="3" Text="" Foreground="#8DA4B5" FontSize="10"/>
        </Grid>
      </TabItem>
      <!-- TAB 5: TOOLS -->
      <TabItem Header="Tools">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#0F1923">
          <StackPanel Margin="14,10">
            <TextBlock Text="App Update" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                  <TextBlock x:Name="TxtCurrentVersion" Text="Current version: ..." Foreground="#8DA4B5" FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
                  <Button x:Name="BtnCheckUpdate" Content="Check for Updates" Background="#1A4A7A" Style="{StaticResource SmallBtn}"/>
                  <Button x:Name="BtnUpdate" Content="Update Now" Background="#1A6B3A" Style="{StaticResource SmallBtn}" Visibility="Collapsed"/>
                  <Button x:Name="BtnPatchNotes" Content="Patch Notes" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
                </StackPanel>
                <TextBlock x:Name="TxtUpdateStatus" Text="" Foreground="#8DA4B5" FontSize="11" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>
            <TextBlock Text="Backup" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                  <Button x:Name="BtnBackup" Content="Backup Saves Now" Background="#1A6B3A" Style="{StaticResource BaseBtn}"/>
                  <Button x:Name="BtnOpenBackups" Content="Open Backup Folder" Background="#1A3A6B" Style="{StaticResource BaseBtn}"/>
                </StackPanel>
                <TextBlock x:Name="TxtLastBackup" Text="Last backup: none" Foreground="#8DA4B5" FontSize="11"/>
                <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                  <CheckBox x:Name="ChkAutoBackup" Content="Auto-backup every"
                            Style="{StaticResource DarkCheck}" VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <ComboBox x:Name="CmbBackupInterval" Width="90" Style="{StaticResource DarkCombo}"
                            VerticalAlignment="Center" SelectedIndex="1">
                    <ComboBoxItem Content="1 hour"/>
                    <ComboBoxItem Content="4 hours"/>
                    <ComboBoxItem Content="8 hours"/>
                    <ComboBoxItem Content="16 hours"/>
                    <ComboBoxItem Content="24 hours"/>
                  </ComboBox>
                </StackPanel>
                <TextBlock x:Name="TxtNextBackup" Text="" Foreground="#8DA4B5" FontSize="11" Margin="0,4,0,0"/>
              </StackPanel>
            </Border>
            <TextBlock Text="Scheduled Restart" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <CheckBox x:Name="ChkSchedule" Content="Enable daily restart at" Style="{StaticResource DarkCheck}" Margin="0,0,10,0"/>
                  <TextBox x:Name="TxtScheduleTime" Text="04:00" Style="{StaticResource DarkInput}" Width="60"/>
                </StackPanel>
              </StackPanel>
            </Border>
            <TextBlock Text="Restart Warning" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="Countdown (seconds):" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtCountdown" Text="30" Style="{StaticResource DarkInput}" Width="60"/>
                </StackPanel>
              </StackPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </TabItem>
      <!-- TAB 6: SETUP WIZARD -->
      <TabItem Header="Install">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#0F1923">
          <StackPanel Margin="14,10">

            <!-- Banner -->
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12,10" Margin="0,0,0,10">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="Server Setup" FontSize="14" FontWeight="Bold" Foreground="#D4A843"/>
                  <TextBlock Text="Follow these steps to get your Windrose dedicated server running." FontSize="11" Foreground="#8DA4B5" Margin="0,3,0,0" TextWrapping="Wrap"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                  <Ellipse x:Name="DotInstall" Width="12" Height="12" Fill="#CC3333" Margin="0,0,6,0" VerticalAlignment="Center"/>
                  <TextBlock x:Name="TxtInstallStatus" Text="Not installed" Foreground="#CC3333" FontSize="11" VerticalAlignment="Center"/>
                </StackPanel>
              </Grid>
            </Border>

            <!-- STEP 1: REQUIREMENTS -->
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,6">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Border x:Name="StepBadge1" Width="26" Height="26" CornerRadius="13" Background="#2A3E55" VerticalAlignment="Center" Margin="0,0,10,0">
                    <TextBlock x:Name="StepBadgeTxt1" Text="1" Foreground="White" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <TextBlock Grid.Column="1" Text="Check Requirements" FontSize="13" FontWeight="Bold" Foreground="#C0CDD8" VerticalAlignment="Center"/>
                  <TextBlock x:Name="StepStatus1" Grid.Column="2" Text="" FontSize="11" Foreground="#8DA4B5" VerticalAlignment="Center"/>
                </Grid>
                <StackPanel Margin="36,0,0,0">
                  <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#8DA4B5" Margin="0,0,0,8">Windrose must be installed via Steam (App ID 3041230). The dedicated server files are bundled inside the game - no separate download needed.</TextBlock>
                  <TextBlock x:Name="TxtReqSteam" Text="Checking..." FontSize="11" Foreground="#8DA4B5" Margin="0,0,0,8"/>
                  <Button x:Name="BtnCheckReqs" Content="Re-check" Background="#2A3E55" Style="{StaticResource SmallBtn}" HorizontalAlignment="Left"/>
                </StackPanel>
              </StackPanel>
            </Border>

            <!-- STEP 2: INSTALL SERVER FILES -->
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,6">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Border x:Name="StepBadge2" Width="26" Height="26" CornerRadius="13" Background="#2A3E55" VerticalAlignment="Center" Margin="0,0,10,0">
                    <TextBlock x:Name="StepBadgeTxt2" Text="2" Foreground="White" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <TextBlock Grid.Column="1" Text="Install Server Files" FontSize="13" FontWeight="Bold" Foreground="#C0CDD8" VerticalAlignment="Center"/>
                  <TextBlock x:Name="StepStatus2" Grid.Column="2" Text="" FontSize="11" Foreground="#8DA4B5" VerticalAlignment="Center"/>
                </Grid>
                <StackPanel Margin="36,0,0,0">
                  <TextBlock Text="Steam Source" Style="{StaticResource FieldLabel}" Margin="0,0,0,4"/>
                  <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtSteamSource" Grid.Column="0" Style="{StaticResource DarkInput}" Margin="0,0,4,0"/>
                    <Button x:Name="BtnDetectSteam" Grid.Column="1" Content="Auto-Detect" Background="#1A4A7A" Style="{StaticResource SmallBtn}" Margin="0,0,4,0"/>
                    <Button x:Name="BtnBrowseSource" Grid.Column="2" Content="Browse..." Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
                  </Grid>
                  <TextBlock Text="Install Destination" Style="{StaticResource FieldLabel}" Margin="0,0,0,4"/>
                  <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtInstallDest" Grid.Column="0" Style="{StaticResource DarkInput}" Margin="0,0,4,0"/>
                    <Button x:Name="BtnBrowseDest" Grid.Column="1" Content="Browse..." Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
                  </Grid>
                  <ScrollViewer Height="110" VerticalScrollBarVisibility="Auto" Margin="0,0,0,8">
                    <TextBox x:Name="TxtInstallLog" IsReadOnly="True" FontFamily="Consolas" FontSize="10" Background="#0A1218" Foreground="#90A8B8" BorderBrush="#1E3348" BorderThickness="1" TextWrapping="Wrap"/>
                  </ScrollViewer>
                  <Button x:Name="BtnInstall" Content="Install Server" Background="#1A6B3A" Style="{StaticResource BaseBtn}" HorizontalAlignment="Left"/>
                </StackPanel>
              </StackPanel>
            </Border>

            <!-- STEP 3: CONFIGURE SERVER -->
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,6">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Border x:Name="StepBadge3" Width="26" Height="26" CornerRadius="13" Background="#2A3E55" VerticalAlignment="Center" Margin="0,0,10,0">
                    <TextBlock x:Name="StepBadgeTxt3" Text="3" Foreground="White" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <TextBlock Grid.Column="1" Text="Name Your Server" FontSize="13" FontWeight="Bold" Foreground="#C0CDD8" VerticalAlignment="Center"/>
                  <TextBlock x:Name="StepStatus3" Grid.Column="2" Text="" FontSize="11" Foreground="#8DA4B5" VerticalAlignment="Center"/>
                </Grid>
                <StackPanel Margin="36,0,0,0">
                  <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#8DA4B5" Margin="0,0,0,10">Set a name and basic options for your server. You can always change these later in the Config tab.</TextBlock>
                  <Grid Margin="0,0,0,6">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="120"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Server Name" Style="{StaticResource FieldLabel}" VerticalAlignment="Center"/>
                    <TextBox x:Name="TxtSetupName" Grid.Column="1" Style="{StaticResource DarkInput}" Text="My Windrose Server"/>
                  </Grid>
                  <Grid Margin="0,0,0,6">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="120"/>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="36"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Max Players" Style="{StaticResource FieldLabel}" VerticalAlignment="Center"/>
                    <Slider x:Name="SlSetupMaxPlayers" Grid.Column="1" Minimum="1" Maximum="20" Value="10" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                    <TextBlock x:Name="TxtSetupMaxVal" Grid.Column="2" Text="10" Foreground="#D4A843" FontWeight="Bold" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                  </Grid>
                  <Grid Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="120"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Password" Style="{StaticResource FieldLabel}" VerticalAlignment="Center"/>
                    <CheckBox x:Name="ChkSetupPassword" Grid.Column="1" Style="{StaticResource DarkCheck}" Margin="0,0,8,0"/>
                    <TextBox x:Name="TxtSetupPassword" Grid.Column="2" Style="{StaticResource DarkInput}" IsEnabled="False"/>
                  </Grid>
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="BtnSaveSetup" Content="Save &amp; Continue" Background="#1A6B3A" Style="{StaticResource BaseBtn}"/>
                    <TextBlock x:Name="TxtSetupStatus" Text="" Foreground="#8DA4B5" FontSize="11" VerticalAlignment="Center" Margin="10,0,0,0"/>
                  </StackPanel>
                </StackPanel>
              </StackPanel>
            </Border>

            <!-- STEP 4: PORT FORWARDING -->
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,6">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Border x:Name="StepBadge4" Width="26" Height="26" CornerRadius="13" Background="#2A3E55" VerticalAlignment="Center" Margin="0,0,10,0">
                    <TextBlock x:Name="StepBadgeTxt4" Text="4" Foreground="White" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <TextBlock Grid.Column="1" Text="Open Your Ports" FontSize="13" FontWeight="Bold" Foreground="#C0CDD8" VerticalAlignment="Center"/>
                  <TextBlock x:Name="StepStatus4" Grid.Column="2" Text="" FontSize="11" Foreground="#8DA4B5" VerticalAlignment="Center"/>
                </Grid>
                <StackPanel Margin="36,0,0,0">
                  <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#8DA4B5" Margin="0,0,0,8">For players outside your home network to connect, log into your router and forward these two ports to this PC:</TextBlock>
                  <Border Background="#0D1820" CornerRadius="4" Padding="10,8" Margin="0,0,0,8">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                        <TextBlock Text="UDP  7777" FontFamily="Consolas" FontSize="12" Foreground="#D4A843" Width="120"/>
                        <TextBlock Text="Game traffic" FontSize="11" Foreground="#8DA4B5" VerticalAlignment="Center"/>
                      </StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="UDP  7778" FontFamily="Consolas" FontSize="12" Foreground="#D4A843" Width="120"/>
                        <TextBlock Text="Game traffic (secondary)" FontSize="11" Foreground="#8DA4B5" VerticalAlignment="Center"/>
                      </StackPanel>
                    </StackPanel>
                  </Border>
                  <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#607080" Margin="0,0,0,8">Players on the same Wi-Fi network as you can connect without port forwarding. Only needed for friends connecting over the internet.</TextBlock>
                  <Button x:Name="BtnCopyPorts" Content="Copy Ports to Clipboard" Background="#2A3E55" Style="{StaticResource SmallBtn}" HorizontalAlignment="Left"/>
                </StackPanel>
              </StackPanel>
            </Border>

            <!-- STEP 5: ALL DONE -->
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,6">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Border x:Name="StepBadge5" Width="26" Height="26" CornerRadius="13" Background="#2A3E55" VerticalAlignment="Center" Margin="0,0,10,0">
                    <TextBlock x:Name="StepBadgeTxt5" Text="5" Foreground="White" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <TextBlock Grid.Column="1" Text="Start Your Server" FontSize="13" FontWeight="Bold" Foreground="#C0CDD8" VerticalAlignment="Center"/>
                </Grid>
                <StackPanel Margin="36,0,0,0">
                  <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#8DA4B5" Margin="0,0,0,10">Everything is ready. Head to the Dashboard tab and click Start. Your invite code will appear in the header once the server is online - share it with friends to let them join.</TextBlock>
                  <Button x:Name="BtnGoToDashboard" Content="Go to Dashboard" Background="#1A6B3A" Style="{StaticResource BaseBtn}" HorizontalAlignment="Left"/>
                </StackPanel>
              </StackPanel>
            </Border>

          </StackPanel>
        </ScrollViewer>
      </TabItem>
    </TabControl>
    <!-- FOOTER BUTTONS -->
    <UniformGrid Grid.Row="2" Rows="1" Columns="5" Margin="8,4">
      <Button x:Name="BtnStart" Content="Start" Background="#1A6B3A" Style="{StaticResource BaseBtn}"/>
      <Button x:Name="BtnStop" Content="Stop" Background="#6B1A1A" Style="{StaticResource BaseBtn}" IsEnabled="False"/>
      <Button x:Name="BtnRestart" Content="Restart" Background="#1A3A7A" Style="{StaticResource BaseBtn}" IsEnabled="False"/>
      <Button x:Name="BtnSave" Content="Save World" Background="#2A3A4A" Style="{StaticResource BaseBtn}" IsEnabled="False"/>
      <Button x:Name="BtnFolder" Content="Open Folder" Background="#1A4A2A" Style="{StaticResource BaseBtn}"/>
    </UniformGrid>
    <!-- STATUS BAR -->
    <Border Grid.Row="3" Background="#0A1218" Padding="10,6">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtLog" Grid.Column="0" Text="Ready." Foreground="#607080" FontSize="11" VerticalAlignment="Center"/>
        <Button x:Name="BtnCancelRestart" Grid.Column="1" Content="Cancel Restart"
                Background="#7A3A1A" Style="{StaticResource SmallBtn}" Visibility="Collapsed"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$Reader = [System.Xml.XmlNodeReader]::new($Xaml)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Override the system brush that WPF's default ComboBox popup uses as its
# background. Must be done in code after the window is created because
# {x:Static SystemColors.*Key} as x:Key in a Style.Setter.Value crashes
# WPF's XamlReader when loaded from a heredoc in PowerShell.
$Window.Resources[[System.Windows.SystemColors]::WindowBrushKey] =
    [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x27,0x36))
$Window.Resources[[System.Windows.SystemColors]::WindowTextBrushKey] =
    [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD0,0xD8,0xE4))
$Window.Resources[[System.Windows.SystemColors]::HighlightBrushKey] =
    [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1E,0x33,0x48))
$Window.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey] =
    [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD4,0xA8,0x43))

function Ctrl($n) { $Window.FindName($n) }

$TxtServerName   = Ctrl 'TxtServerName'
$DotStatus       = Ctrl 'DotStatus'
$TxtStatus       = Ctrl 'TxtStatus'
$TxtUptime       = Ctrl 'TxtUptime'
$TxtInvite       = Ctrl 'TxtInvite'
$BtnShare        = Ctrl 'BtnShare'
$TxtCpu          = Ctrl 'TxtCpu'
$TxtRam          = Ctrl 'TxtRam'
$TxtPlayers      = Ctrl 'TxtPlayers'
$TxtUptimeBig    = Ctrl 'TxtUptimeBig'
$PlayerList          = Ctrl 'PlayerList'
$BtnRefreshPlayers   = Ctrl 'BtnRefreshPlayers'
$ChkAutoRestart      = Ctrl 'ChkAutoRestart'
$ChkSaveOnStop       = Ctrl 'ChkSaveOnStop'
$CfgName         = Ctrl 'CfgName'
$CfgMaxPlayers   = Ctrl 'CfgMaxPlayers'
$TxtMaxPlayersVal= Ctrl 'TxtMaxPlayersVal'
$CfgPasswordEnabled = Ctrl 'CfgPasswordEnabled'
$CfgPassword     = Ctrl 'CfgPassword'
$CfgProxy        = Ctrl 'CfgProxy'
$CfgPreset       = Ctrl 'CfgPreset'
$PanelCustom     = Ctrl 'PanelCustom'
$SlMobHealth     = Ctrl 'SlMobHealth'
$SlMobDamage     = Ctrl 'SlMobDamage'
$SlShipHealth    = Ctrl 'SlShipHealth'
$SlShipDamage    = Ctrl 'SlShipDamage'
$SlBoarding      = Ctrl 'SlBoarding'
$SlCoopStats     = Ctrl 'SlCoopStats'
$SlCoopShip      = Ctrl 'SlCoopShip'
$ValMobHealth    = Ctrl 'ValMobHealth'
$ValMobDamage    = Ctrl 'ValMobDamage'
$ValShipHealth   = Ctrl 'ValShipHealth'
$ValShipDamage   = Ctrl 'ValShipDamage'
$ValBoarding     = Ctrl 'ValBoarding'
$ValCoopStats    = Ctrl 'ValCoopStats'
$ValCoopShip     = Ctrl 'ValCoopShip'
$CfgCombatDiff   = Ctrl 'CfgCombatDiff'
$CfgCoopQuests   = Ctrl 'CfgCoopQuests'
$CfgEasyExplore  = Ctrl 'CfgEasyExplore'
$BtnSaveConfig   = Ctrl 'BtnSaveConfig'
$BtnReloadConfig = Ctrl 'BtnReloadConfig'
$BtnOpenWorldJson= Ctrl 'BtnOpenWorldJson'
$TxtConfigStatus = Ctrl 'TxtConfigStatus'
$BtnFAll         = Ctrl 'BtnFAll'
$BtnFPlayers     = Ctrl 'BtnFPlayers'
$BtnFWarn        = Ctrl 'BtnFWarn'
$BtnFErrors      = Ctrl 'BtnFErrors'
$ChkAutoScroll   = Ctrl 'ChkAutoScroll'
$LogViewer       = Ctrl 'LogViewer'
$ConsoleOutput   = Ctrl 'ConsoleOutput'
$BtnCmdSave      = Ctrl 'BtnCmdSave'
$BtnCmdPlayers   = Ctrl 'BtnCmdPlayers'
$BtnCmdInfo      = Ctrl 'BtnCmdInfo'
$BtnCmdQuit      = Ctrl 'BtnCmdQuit'
$TxtCommand      = Ctrl 'TxtCommand'
$BtnSendCmd      = Ctrl 'BtnSendCmd'
$TxtConsoleStatus= Ctrl 'TxtConsoleStatus'
$BtnBackup       = Ctrl 'BtnBackup'
$BtnOpenBackups  = Ctrl 'BtnOpenBackups'
$TxtLastBackup   = Ctrl 'TxtLastBackup'
$ChkAutoBackup   = Ctrl 'ChkAutoBackup'
$CmbBackupInterval = Ctrl 'CmbBackupInterval'
$TxtNextBackup   = Ctrl 'TxtNextBackup'
$ChkSchedule     = Ctrl 'ChkSchedule'
$TxtScheduleTime = Ctrl 'TxtScheduleTime'
$TxtCountdown    = Ctrl 'TxtCountdown'
$HistoryList     = Ctrl 'HistoryList'
$BtnClearHistory = Ctrl 'BtnClearHistory'
$TxtCurrentVersion = Ctrl 'TxtCurrentVersion'
$BtnCheckUpdate  = Ctrl 'BtnCheckUpdate'
$BtnUpdate       = Ctrl 'BtnUpdate'
$BtnPatchNotes   = Ctrl 'BtnPatchNotes'
$TxtUpdateStatus = Ctrl 'TxtUpdateStatus'
$DotInstall      = Ctrl 'DotInstall'
$TxtInstallStatus= Ctrl 'TxtInstallStatus'
$TxtSteamSource  = Ctrl 'TxtSteamSource'
$BtnDetectSteam  = Ctrl 'BtnDetectSteam'
$BtnBrowseSource = Ctrl 'BtnBrowseSource'
$TxtInstallDest  = Ctrl 'TxtInstallDest'
$BtnBrowseDest   = Ctrl 'BtnBrowseDest'
$TxtInstallLog   = Ctrl 'TxtInstallLog'
$BtnInstall      = Ctrl 'BtnInstall'
# Wizard controls
$StepBadge1      = Ctrl 'StepBadge1';  $StepBadgeTxt1 = Ctrl 'StepBadgeTxt1';  $StepStatus1 = Ctrl 'StepStatus1'
$StepBadge2      = Ctrl 'StepBadge2';  $StepBadgeTxt2 = Ctrl 'StepBadgeTxt2';  $StepStatus2 = Ctrl 'StepStatus2'
$StepBadge3      = Ctrl 'StepBadge3';  $StepBadgeTxt3 = Ctrl 'StepBadgeTxt3';  $StepStatus3 = Ctrl 'StepStatus3'
$StepBadge4      = Ctrl 'StepBadge4';  $StepBadgeTxt4 = Ctrl 'StepBadgeTxt4';  $StepStatus4 = Ctrl 'StepStatus4'
$StepBadge5      = Ctrl 'StepBadge5';  $StepBadgeTxt5 = Ctrl 'StepBadgeTxt5'
$TxtReqSteam     = Ctrl 'TxtReqSteam'
$BtnCheckReqs    = Ctrl 'BtnCheckReqs'
$TxtSetupName    = Ctrl 'TxtSetupName'
$SlSetupMaxPlayers = Ctrl 'SlSetupMaxPlayers'
$TxtSetupMaxVal  = Ctrl 'TxtSetupMaxVal'
$ChkSetupPassword = Ctrl 'ChkSetupPassword'
$TxtSetupPassword = Ctrl 'TxtSetupPassword'
$BtnSaveSetup    = Ctrl 'BtnSaveSetup'
$TxtSetupStatus  = Ctrl 'TxtSetupStatus'
$BtnCopyPorts    = Ctrl 'BtnCopyPorts'
$BtnGoToDashboard = Ctrl 'BtnGoToDashboard'
$BtnStart        = Ctrl 'BtnStart'
$BtnStop         = Ctrl 'BtnStop'
$BtnRestart      = Ctrl 'BtnRestart'
$BtnSave         = Ctrl 'BtnSave'
$BtnFolder       = Ctrl 'BtnFolder'
$TxtLog          = Ctrl 'TxtLog'
$BtnCancelRestart= Ctrl 'BtnCancelRestart'
$MainTabs        = Ctrl 'MainTabs'

# Script-scoped state
$script:ServerProc      = $null
$script:StartTime       = $null
$script:PrevCpuTime     = $null
$script:PrevCpuCheck    = $null
$script:MaxPlayers      = 10
$script:pollTimer       = $null
$script:countdownTimer  = $null
$script:countdownSecs   = 0
$script:countdownAction = $null
$script:logBuffer       = [System.Collections.Generic.List[string]]::new()
$script:logPosition     = 0L
$script:onlinePlayers   = [System.Collections.Generic.HashSet[string]]::new()
$script:logFilter       = "All"
$script:scheduleFired   = $false
$script:lastScheduleDate= $null
$script:consoleBuffer   = [System.Collections.Generic.List[string]]::new()
$script:installCopyJob  = $null
$script:installTimer    = $null
$script:updateCheckJob  = $null
$script:updatePollTimer = $null
$script:lastPlayerSnapshot = ""
$script:accountToPlayer    = @{}
$script:autoBackupTimer    = $null
$script:lastBackupStamp    = $null

# Cached brushes & fonts -- avoids re-allocating identical objects every tick / every log line
$script:FontConsolas    = [System.Windows.Media.FontFamily]::new("Consolas")
$script:BrushGrayBtn    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55)); $script:BrushGrayBtn.Freeze()
$script:BrushBlueBtn    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x4A,0x7A)); $script:BrushBlueBtn.Freeze()
$script:BrushGrayText   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x8D,0xA4,0xB5)); $script:BrushGrayText.Freeze()
$script:BrushStopped    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x55,0x55,0x55)); $script:BrushStopped.Freeze()
$script:BrushLogDefault = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x70,0x88,0x99)); $script:BrushLogDefault.Freeze()
$script:BrushLeave      = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xFA,0x80,0x72)); $script:BrushLeave.Freeze()
$script:BrushGoldCmd    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD4,0xA8,0x43)); $script:BrushGoldCmd.Freeze()
$script:BrushRed        = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xCC,0x33,0x33)); $script:BrushRed.Freeze()

# ---- HELPERS ----

function Get-ServerProcess {
    try {
        $p = Get-Process "WindroseServer-Win64-Shipping" -ErrorAction SilentlyContinue
        if ($p) { return $p | Select-Object -First 1 }
    } catch {}
    try {
        $p = Get-Process "WindroseServer" -ErrorAction SilentlyContinue
        if ($p) { return $p | Select-Object -First 1 }
    } catch {}
    return $null
}

function Stop-AllServerProcesses {
    try { Get-Process "WindroseServer*" -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() } } catch {}
}

function Read-InviteCode {
    try {
        if (Test-Path $ConfigPath) {
            $j = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            $inner = $j.ServerDescription_Persistent
            if ($inner -and $inner.PSObject.Properties['InviteCode']) { return $inner.InviteCode }
        }
    } catch {}
    return $null
}

function Read-ServerConfig {
    try {
        if (-not (Test-Path $ConfigPath)) { return }
        $j = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $inner = $j.ServerDescription_Persistent
        if (-not $inner) { return }
        if ($inner.PSObject.Properties['ServerName']) {
            $CfgName.Text = $inner.ServerName
            $TxtServerName.Text = $inner.ServerName
        }
        if ($inner.PSObject.Properties['MaxPlayerCount']) {
            $val = [double]$inner.MaxPlayerCount
            if ($val -lt 1)  { $val = 1  }
            if ($val -gt 20) { $val = 20 }
            $CfgMaxPlayers.Value = $val
            $script:MaxPlayers = [int]$val
            $TxtMaxPlayersVal.Text = "$([int]$val)"
        }
        $isProtected = $false
        if ($inner.PSObject.Properties['IsPasswordProtected']) { $isProtected = [bool]$inner.IsPasswordProtected }
        if ($isProtected) {
            $CfgPasswordEnabled.IsChecked = $true
            $CfgPassword.IsEnabled = $true
            if ($inner.PSObject.Properties['Password']) { $CfgPassword.Text = $inner.Password }
        } else {
            $CfgPasswordEnabled.IsChecked = $false
            $CfgPassword.IsEnabled = $false
        }
        if ($inner.PSObject.Properties['P2pProxyAddress']) { $CfgProxy.Text = $inner.P2pProxyAddress }
    } catch {}
}

function Find-WorldConfig {
    $rockPath = "$ServerDir\R5\Saved\SaveProfiles\Default\RocksDB"
    if (-not (Test-Path $rockPath)) { return $null }
    $files = Get-ChildItem -Path $rockPath -Recurse -Filter "WorldDescription.json" -ErrorAction SilentlyContinue
    if (-not $files) { return $null }
    return ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

function Read-WorldConfig {
    try {
        $wPath = Find-WorldConfig
        if (-not $wPath) { return }
        $j = Get-Content $wPath -Raw | ConvertFrom-Json

        $wd = $j.WorldDescription
        if (-not $wd) { return }

        $preset = "Custom"
        if ($wd.PSObject.Properties['WorldPresetType']) { $preset = $wd.WorldPresetType }

        $matchedItem = $null
        foreach ($item in $CfgPreset.Items) {
            if ($item.Content -eq $preset) { $matchedItem = $item; break }
        }
        if ($matchedItem) { $CfgPreset.SelectedItem = $matchedItem }
        if ($preset -eq "Custom") { $PanelCustom.Visibility = "Visible" }

        $ws = $wd.WorldSettings
        if (-not $ws) { return }

        # Float parameters -- keys are JSON strings like '{"TagName": "WDS.Parameter.MobHealthMultiplier"}'
        $fp = $ws.FloatParameters
        if ($fp) {
            $floatMap = [ordered]@{
                '{"TagName": "WDS.Parameter.MobHealthMultiplier"}'              = $SlMobHealth
                '{"TagName": "WDS.Parameter.MobDamageMultiplier"}'              = $SlMobDamage
                '{"TagName": "WDS.Parameter.ShipsHealthMultiplier"}'            = $SlShipHealth
                '{"TagName": "WDS.Parameter.ShipsDamageMultiplier"}'            = $SlShipDamage
                '{"TagName": "WDS.Parameter.BoardingDifficultyMultiplier"}'     = $SlBoarding
                '{"TagName": "WDS.Parameter.Coop.StatsCorrectionModifier"}'    = $SlCoopStats
                '{"TagName": "WDS.Parameter.Coop.ShipStatsCorrectionModifier"}' = $SlCoopShip
            }
            foreach ($key in $floatMap.Keys) {
                $prop = $fp.PSObject.Properties[$key]
                if ($prop) { $floatMap[$key].Value = [double]$prop.Value }
            }
        }

        # Bool parameters
        $bp = $ws.BoolParameters
        if ($bp) {
            $prop = $bp.PSObject.Properties['{"TagName": "WDS.Parameter.Coop.SharedQuests"}']
            if ($prop) { $CfgCoopQuests.IsChecked  = [bool]$prop.Value }
            $prop = $bp.PSObject.Properties['{"TagName": "WDS.Parameter.EasyExplore"}']
            if ($prop) { $CfgEasyExplore.IsChecked = [bool]$prop.Value }
        }

        # Tag parameters -- combat difficulty
        $tp = $ws.TagParameters
        if ($tp) {
            $prop = $tp.PSObject.Properties['{"TagName": "WDS.Parameter.CombatDifficulty"}']
            if ($prop) {
                $tagName = $prop.Value.TagName  # e.g. "WDS.Parameter.CombatDifficulty.Normal"
                $short   = $tagName -replace '^.*\.CombatDifficulty\.', ''  # "Normal"
                foreach ($item in $CfgCombatDiff.Items) {
                    if ($item.Content -eq $short) { $CfgCombatDiff.SelectedItem = $item; break }
                }
            }
        }
    } catch {}
}

function Find-SteamWindrose {
    $steamPath = $null
    try {
        $reg = Get-ItemProperty "HKCU:\Software\Valve\Steam" -ErrorAction SilentlyContinue
        if ($reg -and $reg.SteamPath) { $steamPath = $reg.SteamPath }
    } catch {}
    if (-not $steamPath) {
        try {
            $reg = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Valve\Steam" -ErrorAction SilentlyContinue
            if ($reg -and $reg.InstallPath) { $steamPath = $reg.InstallPath }
        } catch {}
    }
    if (-not $steamPath) { $steamPath = "C:\Program Files (x86)\Steam" }

    $libraryPaths = @($steamPath)
    $vdfPath = "$steamPath\steamapps\libraryfolders.vdf"
    if (Test-Path $vdfPath) {
        $content = Get-Content $vdfPath -Raw
        $matches2 = [regex]::Matches($content, '"path"\s+"([^"]+)"')
        foreach ($m in $matches2) {
            $libraryPaths += $m.Groups[1].Value -replace '\\\\', '\'
        }
    }
    foreach ($lib in $libraryPaths) {
        $candidate = "$lib\steamapps\common\Windrose\R5\Builds\WindowsServer"
        if (Test-Path "$candidate\WindroseServer.exe") {
            return $candidate
        }
    }
    return $null
}

function New-LogTextBlock($line) {
    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.Text = $line
    $tb.FontFamily = $script:FontConsolas
    $tb.FontSize = 11
    $tb.TextWrapping = "NoWrap"
    $low = $line.ToLower()
    if     ($low -match 'error|fatal')   { $tb.Foreground = [System.Windows.Media.Brushes]::Tomato }
    elseif ($low -match 'warning')       { $tb.Foreground = [System.Windows.Media.Brushes]::Orange }
    elseif ($low -match 'join succeeded'){ $tb.Foreground = [System.Windows.Media.Brushes]::LightGreen }
    elseif ($low -match 'leave:|saidfarewell|disconnectaccount') { $tb.Foreground = $script:BrushLeave }
    else   { $tb.Foreground = $script:BrushLogDefault }
    return $tb
}

function Test-LogFilter($low, $filter) {
    switch ($filter) {
        "All"     { return $true }
        "Players" { return ($low -match 'join succeeded|leave:|saidfarewell|disconnectaccount') }
        "Warn"    { return ($low -match 'warning') }
        "Errors"  { return ($low -match 'error|fatal') }
    }
    return $true
}

function Update-LogViewer {
    if (-not (Test-Path $LogPath)) { return }
    try {
        $fs = [System.IO.FileStream]::new($LogPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite)
        $fs.Seek($script:logPosition, [System.IO.SeekOrigin]::Begin) | Out-Null
        $reader = [System.IO.StreamReader]::new($fs)
        $newLines = [System.Collections.Generic.List[string]]::new()
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ($line -ne $null) { $newLines.Add($line) }
        }
        $script:logPosition = $fs.Position
        $reader.Dispose()
        $fs.Dispose()

        if ($newLines.Count -eq 0) { return }

        foreach ($line in $newLines) {
            $script:logBuffer.Add($line)
            $low = $line.ToLower()

            # --- Player join ---
            if ($low -match 'lognet: join succeeded:') {
                $playerName = ""
                if ($line -match 'Join succeeded:\s*(.+)') { $playerName = $Matches[1].Trim() }
                if ($playerName) {
                    $script:onlinePlayers.Add($playerName) | Out-Null
                    Add-History "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] JOINED: $playerName"
                }
            }
            # Build AccountId -> PlayerName mapping (appears near join)
            elseif ($line -match "AccountName '([^']+)'.*AccountId (\w+)") {
                $script:accountToPlayer[$Matches[2]] = $Matches[1]
            }
            # --- Player leave: graceful (standard UE5) ---
            elseif ($low -match 'lognet: leave:') {
                $playerName = ""
                if ($line -match 'Leave:\s*(.+)') { $playerName = $Matches[1].Trim() }
                if ($playerName) {
                    $script:onlinePlayers.Remove($playerName) | Out-Null
                    Add-History "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] LEFT: $playerName"
                }
            }
            # --- Player leave: Windrose farewell (lobby quit, graceful) ---
            elseif ($line -match "Name '([^']+)'.*State 'SaidFarewell'") {
                $playerName = $Matches[1]
                if ($script:onlinePlayers.Remove($playerName)) {
                    Add-History "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] LEFT: $playerName"
                }
            }
            # --- Player leave: DisconnectAccount (crash, timeout, any disconnect) ---
            elseif ($low -match 'disconnectaccount.*accountid (\w+)') {
                $acctId = $Matches[1]
                $playerName = $script:accountToPlayer[$acctId]
                if ($playerName -and $script:onlinePlayers.Remove($playerName)) {
                    Add-History "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] LEFT: $playerName (disconnect)"
                }
            }
        }

        # Cap buffer
        if ($script:logBuffer.Count -gt 1000) {
            $excess = $script:logBuffer.Count - 1000
            $script:logBuffer.RemoveRange(0, $excess)
        }

        # Feed new lines to Console tab so command responses are visible
        foreach ($line in $newLines) {
            Add-ConsoleEntry $line
        }

        # Append only new lines to LogViewer (instead of full rebuild)
        $filter = $script:logFilter
        foreach ($line in $newLines) {
            if (Test-LogFilter $line.ToLower() $filter) {
                $LogViewer.Items.Add((New-LogTextBlock $line)) | Out-Null
            }
        }

        # Cap viewer items to match buffer
        while ($LogViewer.Items.Count -gt 1000) {
            $LogViewer.Items.RemoveAt(0)
        }

        if ($ChkAutoScroll.IsChecked -and $LogViewer.Items.Count -gt 0) {
            $LogViewer.ScrollIntoView($LogViewer.Items[$LogViewer.Items.Count - 1])
        }
    } catch {}
}

# Full rebuild -- only called when the log filter changes
function Refresh-LogViewer {
    $LogViewer.Items.Clear()
    $filter = $script:logFilter
    foreach ($line in $script:logBuffer) {
        if (Test-LogFilter $line.ToLower() $filter) {
            $LogViewer.Items.Add((New-LogTextBlock $line)) | Out-Null
        }
    }
    if ($ChkAutoScroll.IsChecked -and $LogViewer.Items.Count -gt 0) {
        $LogViewer.ScrollIntoView($LogViewer.Items[$LogViewer.Items.Count - 1])
    }
}

function Add-ConsoleEntry($text, $isCommand = $false) {
    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.TextWrapping = "NoWrap"
    $tb.FontFamily = $script:FontConsolas
    $tb.FontSize = 11
    if ($isCommand) {
        $tb.Foreground = $script:BrushGoldCmd
        $tb.Text = "> $text"
    } else {
        $low = $text.ToLower()
        if     ($low -match 'error|fatal')   { $tb.Foreground = [System.Windows.Media.Brushes]::Tomato }
        elseif ($low -match 'warning')       { $tb.Foreground = [System.Windows.Media.Brushes]::Orange }
        elseif ($low -match 'join succeeded'){ $tb.Foreground = [System.Windows.Media.Brushes]::LightGreen }
        elseif ($low -match 'leave:|saidfarewell|disconnectaccount') { $tb.Foreground = $script:BrushLeave }
        else   { $tb.Foreground = $script:BrushLogDefault }
        $tb.Text = $text
    }
    $ConsoleOutput.Items.Add($tb) | Out-Null
    if ($ConsoleOutput.Items.Count -gt 500) {
        $ConsoleOutput.Items.RemoveAt(0)
    }
    if ($ConsoleOutput.Items.Count -gt 0) {
        $ConsoleOutput.ScrollIntoView($ConsoleOutput.Items[$ConsoleOutput.Items.Count - 1])
    }
}

function Read-PlayerList {
    # Rebuilds $script:onlinePlayers by replaying the full log.
    # Uses FileShare.ReadWrite so it works while the server has the log open.
    $online = [System.Collections.Generic.HashSet[string]]::new()
    $acctMap = @{}
    if (-not (Test-Path $LogPath)) { return @() }
    try {
        $fs = [System.IO.FileStream]::new($LogPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite)
        $reader = [System.IO.StreamReader]::new($fs)
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ($null -eq $line) { continue }
            $low = $line.ToLower()
            # Join
            if ($low -match 'lognet: join succeeded:') {
                if ($line -match 'Join succeeded:\s*(.+)') { $online.Add($Matches[1].Trim()) | Out-Null }
            }
            # AccountId mapping
            elseif ($line -match "AccountName '([^']+)'.*AccountId (\w+)") {
                $acctMap[$Matches[2]] = $Matches[1]
            }
            # Standard UE5 leave
            elseif ($low -match 'lognet: leave:') {
                if ($line -match 'Leave:\s*(.+)') { $online.Remove($Matches[1].Trim()) | Out-Null }
            }
            # Windrose farewell
            elseif ($line -match "Name '([^']+)'.*State 'SaidFarewell'") {
                $online.Remove($Matches[1]) | Out-Null
            }
            # DisconnectAccount (crash, timeout, any disconnect)
            elseif ($low -match 'disconnectaccount.*accountid (\w+)') {
                $pn = $acctMap[$Matches[1]]
                if ($pn) { $online.Remove($pn) | Out-Null }
            }
        }
        $reader.Dispose()
        $fs.Dispose()
    } catch {}
    $script:onlinePlayers.Clear()
    $script:accountToPlayer = $acctMap
    foreach ($p in $online) { $script:onlinePlayers.Add($p) | Out-Null }
    return @($online)
}

function Add-History($entry) {
    $HistoryList.Items.Add($entry) | Out-Null
    try { Add-Content -Path $HistoryFile -Value $entry -Encoding UTF8 } catch {}
}

function Load-History {
    if (Test-Path $HistoryFile) {
        try {
            $lines = Get-Content $HistoryFile -Tail 100
            foreach ($l in $lines) { $HistoryList.Items.Add($l) | Out-Null }
        } catch {}
    }
}

function Update-Stats {
    $proc = Get-ServerProcess
    if (-not $proc) { Reset-Stats; return }
    try {
        $now = [DateTime]::Now
        $cpuNow = $proc.TotalProcessorTime
        $cpuPct = 0
        if ($script:PrevCpuTime -ne $null -and $script:PrevCpuCheck -ne $null) {
            $elapsed = ($now - $script:PrevCpuCheck).TotalSeconds
            if ($elapsed -gt 0) {
                $delta = ($cpuNow - $script:PrevCpuTime).TotalSeconds
                $cpuPct = [Math]::Round(($delta / $elapsed / [Environment]::ProcessorCount) * 100, 1)
            }
        }
        $script:PrevCpuTime  = $cpuNow
        $script:PrevCpuCheck = $now
        $TxtCpu.Text = "$cpuPct%"

        $ramMb = [Math]::Round($proc.WorkingSet64 / 1MB, 1)
        if ($ramMb -ge 1024) {
            $TxtRam.Text = "$([Math]::Round($ramMb/1024,1)) GB"
        } else {
            $TxtRam.Text = "$ramMb MB"
        }

        # Only rebuild PlayerList when the player set actually changes
        $snapshot = ($script:onlinePlayers | Sort-Object) -join ','
        if ($snapshot -ne $script:lastPlayerSnapshot) {
            $script:lastPlayerSnapshot = $snapshot
            $PlayerList.Items.Clear()
            foreach ($p in $script:onlinePlayers) { $PlayerList.Items.Add($p) | Out-Null }
        }
        $TxtPlayers.Text = "$($script:onlinePlayers.Count) / $($script:MaxPlayers)"

        if ($script:StartTime) {
            $up = [DateTime]::Now - $script:StartTime
            $upStr = ""
            if ($up.TotalHours -ge 1) { $upStr = "$([int]$up.TotalHours)h $($up.Minutes)m" }
            else { $upStr = "$($up.Minutes)m $($up.Seconds)s" }
            $TxtUptimeBig.Text = $upStr
            $TxtUptime.Text    = "Up: $upStr"
        }
    } catch {}
}

function Refresh-PlayerList {
    # Full log replay -- rebuilds player list from scratch, correcting any missed leave events.
    Read-PlayerList | Out-Null
    $PlayerList.Items.Clear()
    foreach ($p in $script:onlinePlayers) { $PlayerList.Items.Add($p) | Out-Null }
    $TxtPlayers.Text = "$($script:onlinePlayers.Count) / $($script:MaxPlayers)"
}

function Reset-Stats {
    $TxtCpu.Text = "--"
    $TxtRam.Text = "--"
    $TxtPlayers.Text = "--"
    $TxtUptimeBig.Text = "--"
    $TxtUptime.Text = ""
    $PlayerList.Items.Clear()
}

function Update-SetupWizard {
    $cGreen  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x6B,0x3A))
    $cBlue   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x4A,0x7A))
    $cGray   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $fGreen  = [System.Windows.Media.Brushes]::LightGreen
    $fGray   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x8D,0xA4,0xB5))
    $fRed    = [System.Windows.Media.Brushes]::Tomato

    $steamFound  = ($null -ne (Find-SteamWindrose)) -or (Test-Path $ServerExe)
    $serverReady = Test-Path $ServerExe
    $configReady = ($script:step3Saved -eq $true) -or (Test-Path $ConfigPath)

    # Step 1 -- Requirements
    if ($steamFound) {
        $StepBadge1.Background = $cGreen; $StepBadgeTxt1.Text = [char]0x2713
        $StepStatus1.Text = "Ready"; $StepStatus1.Foreground = $fGreen
        $TxtReqSteam.Text = ([char]0x2713) + " Windrose found on Steam"
        $TxtReqSteam.Foreground = $fGreen
    } else {
        $StepBadge1.Background = $cBlue; $StepBadgeTxt1.Text = "1"
        $StepStatus1.Text = "Action needed"; $StepStatus1.Foreground = $fRed
        $TxtReqSteam.Text = ([char]0x2717) + " Windrose not found - install it via Steam first (App ID 3041230)"
        $TxtReqSteam.Foreground = $fRed
    }

    # Step 2 -- Install
    if ($serverReady) {
        $StepBadge2.Background = $cGreen; $StepBadgeTxt2.Text = [char]0x2713
        $StepStatus2.Text = "Installed"; $StepStatus2.Foreground = $fGreen
    } elseif ($steamFound) {
        $StepBadge2.Background = $cBlue; $StepBadgeTxt2.Text = "2"
        $StepStatus2.Text = "Ready to install"; $StepStatus2.Foreground = $fGray
    } else {
        $StepBadge2.Background = $cGray; $StepBadgeTxt2.Text = "2"
        $StepStatus2.Text = "Complete step 1 first"; $StepStatus2.Foreground = $fGray
    }

    # Step 3 -- Configure
    if ($configReady -and $serverReady) {
        $StepBadge3.Background = $cGreen; $StepBadgeTxt3.Text = [char]0x2713
        $StepStatus3.Text = "Configured"; $StepStatus3.Foreground = $fGreen
    } elseif ($serverReady) {
        $StepBadge3.Background = $cBlue; $StepBadgeTxt3.Text = "3"
        $StepStatus3.Text = "Fill in and save"; $StepStatus3.Foreground = $fGray
    } else {
        $StepBadge3.Background = $cGray; $StepBadgeTxt3.Text = "3"
        $StepStatus3.Text = "Complete step 2 first"; $StepStatus3.Foreground = $fGray
    }

    # Step 4 -- Port forwarding (informational)
    if ($serverReady) {
        $StepBadge4.Background = $cBlue; $StepBadgeTxt4.Text = "4"
        $StepStatus4.Text = "Review ports"; $StepStatus4.Foreground = $fGray
    } else {
        $StepBadge4.Background = $cGray; $StepBadgeTxt4.Text = "4"
        $StepStatus4.Text = ""; $StepStatus4.Foreground = $fGray
    }

    # Step 5 -- Go to dashboard
    if ($serverReady) {
        $StepBadge5.Background = $cBlue; $StepBadgeTxt5.Text = "5"
    } else {
        $StepBadge5.Background = $cGray; $StepBadgeTxt5.Text = "5"
    }
}

function Set-UIRunning {
    $DotStatus.Fill = [System.Windows.Media.Brushes]::LimeGreen
    $TxtStatus.Text = "  Running"
    $TxtStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
    $BtnStart.IsEnabled   = $false
    $BtnStop.IsEnabled    = $true
    $BtnRestart.IsEnabled = $true
    $BtnSave.IsEnabled    = $true
    $code = Read-InviteCode
    if ($code) { $TxtInvite.Text = $code } else { $TxtInvite.Text = "(pending...)" }
}

function Set-UIStopped {
    $DotStatus.Fill = $script:BrushStopped
    $TxtStatus.Text = "  Stopped"
    $TxtStatus.Foreground = $script:BrushGrayText
    $BtnStart.IsEnabled   = $true
    $BtnStop.IsEnabled    = $false
    $BtnRestart.IsEnabled = $false
    $BtnSave.IsEnabled    = $false
    $TxtInvite.Text = "--"
    Reset-Stats
}

function Log($msg) { $TxtLog.Text = $msg }

function Invoke-RestartWithCountdown($actionBlock) {
    $secs = 0
    try { $secs = [int]$TxtCountdown.Text } catch {}
    if ($secs -le 0) {
        & $actionBlock
        return
    }
    $script:countdownSecs   = $secs
    $script:countdownAction = $actionBlock
    $BtnCancelRestart.Visibility = "Visible"
    $script:countdownTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:countdownTimer.Interval = [TimeSpan]::FromSeconds(1)
    $script:countdownTimer.Add_Tick({
        $script:countdownSecs--
        if ($script:countdownSecs -le 0) {
            $script:countdownTimer.Stop()
            $BtnCancelRestart.Visibility = "Collapsed"
            Log "Restarting..."
            & $script:countdownAction
        } else {
            Log "Restarting in $script:countdownSecs seconds..."
        }
    })
    $script:countdownTimer.Start()
    Log "Restarting in $script:countdownSecs seconds..."
}

function Start-ServerProcess {
    $psi = [Diagnostics.ProcessStartInfo]::new()
    # Launch the shipping exe directly -- avoids cmd.exe child spawned by WindroseServer.exe.
    if (Test-Path $ServerExeDirect) {
        $psi.FileName = $ServerExeDirect
    } else {
        $psi.FileName = $ServerExe
    }
    # -log causes UE5 to call AllocConsole() which enables console command processing.
    # We send commands via Win32 WriteConsoleInput (NOT stdin redirect, since
    # AllocConsole overwrites the stdin handle making the .NET pipe useless).
    $psi.Arguments        = "-log"
    $psi.WorkingDirectory = $ServerDir
    $psi.UseShellExecute  = $false
    $psi.CreateNoWindow   = $true
    $script:ServerProc = [Diagnostics.Process]::Start($psi)

    # UE5 calls AllocConsole() during startup which creates a visible window.
    # Poll for 12 seconds and hide any window owned by the server process.
    $procId = $script:ServerProc.Id
    $script:hideAttempts = 0
    $hideTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $hideTimer.Interval = [TimeSpan]::FromSeconds(1)
    $hideTimer.Add_Tick({
        $script:hideAttempts++
        try { [WinHelper]::HideProcessWindows($procId) } catch {}
        if ($script:hideAttempts -ge 12) { $hideTimer.Stop() }
    })
    $hideTimer.Start()
}

function Send-ServerCommand($cmd) {
    $proc = Get-ServerProcess
    if (-not $proc) { Log "Server not running."; return $false }
    try {
        $ok = [WinHelper]::SendConsoleCommand($proc.Id, $cmd)
        if (-not $ok) { Log "Failed to send command (could not attach to console)."; return $false }
        return $true
    } catch { Log "Command error: $_"; return $false }
}

$script:doRestart = {
    Stop-AllServerProcesses
    Start-Sleep -Milliseconds 1500
    Start-ServerProcess
    $script:StartTime      = [DateTime]::Now
    $script:logPosition    = 0L
    $script:logBuffer.Clear()
    $script:onlinePlayers.Clear()
    $script:PrevCpuTime    = $script:ServerProc.TotalProcessorTime
    $script:PrevCpuCheck   = [DateTime]::Now
    Set-UIRunning
    Add-ConsoleEntry "Server restarted."
    Log "Server restarted."
    if ($script:pollTimer -ne $null) { $script:pollTimer.Stop() }
    $script:pollTimer = $null
    $script:pollCount = 0
    $script:pollTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:pollTimer.Interval = [TimeSpan]::FromSeconds(5)
    $script:pollTimer.Add_Tick({
        $script:pollCount++
        $code = Read-InviteCode
        if ($code -and $code -ne "") {
            $TxtInvite.Text = $code
            $script:pollTimer.Stop()
        } elseif ($script:pollCount -ge 24) {
            $script:pollTimer.Stop()
        }
    })
    $script:pollTimer.Start()
}

# ---- EVENT HANDLERS ----

$BtnStart.Add_Click({
    if (-not (Test-Path $ServerExe)) { Log "Server not installed."; return }
    try {
        $script:logPosition = 0L
        $script:logBuffer.Clear()
        $script:onlinePlayers.Clear()
        Start-ServerProcess
        $script:StartTime    = [DateTime]::Now
        $script:PrevCpuTime  = $script:ServerProc.TotalProcessorTime
        $script:PrevCpuCheck = [DateTime]::Now
        Set-UIRunning
        Add-ConsoleEntry "Server starting..."
        Log "Server started."
        if ($script:pollTimer -ne $null) { $script:pollTimer.Stop() }
        $script:pollTimer = $null
        $script:pollCount = 0
        $script:pollTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:pollTimer.Interval = [TimeSpan]::FromSeconds(5)
        $script:pollTimer.Add_Tick({
            $script:pollCount++
            $code = Read-InviteCode
            if ($code -and $code -ne "") {
                $TxtInvite.Text = $code
                $script:pollTimer.Stop()
            } elseif ($script:pollCount -ge 24) {
                $script:pollTimer.Stop()
            }
        })
        $script:pollTimer.Start()
    } catch { Log "Failed to start: $_" }
})

$BtnStop.Add_Click({
    if ($script:pollTimer)      { $script:pollTimer.Stop();      $script:pollTimer      = $null }
    if ($script:countdownTimer) { $script:countdownTimer.Stop(); $script:countdownTimer = $null }
    if ($ChkSaveOnStop.IsChecked) {
        Log "Saving world before stopping..."
        Send-ServerCommand "save world" | Out-Null
        Start-Sleep -Seconds 3
    }
    Stop-AllServerProcesses
    $script:ServerProc  = $null
    $script:StartTime   = $null
    Set-UIStopped
    $BtnCancelRestart.Visibility = "Collapsed"
    if ($ChkSaveOnStop.IsChecked) { Log "World saved. Server stopped." }
    else { Log "Server stopped." }
})

$BtnRestart.Add_Click({
    Invoke-RestartWithCountdown $script:doRestart
})

$BtnSave.Add_Click({
    if (Send-ServerCommand "save world") {
        Log "Save command sent."
        Add-ConsoleEntry "save world" $true
    }
})

$BtnFolder.Add_Click({ Start-Process explorer.exe $ServerDir })

$BtnCancelRestart.Add_Click({
    if ($script:countdownTimer) { $script:countdownTimer.Stop(); $script:countdownTimer = $null }
    $BtnCancelRestart.Visibility = "Collapsed"
    Log "Restart cancelled."
})

$TxtInvite.Add_MouseLeftButtonUp({
    $code = $TxtInvite.Text
    if ($code -and $code -ne "--" -and $code -ne "(pending...)") {
        [System.Windows.Clipboard]::SetText($code)
        Log "Invite code copied to clipboard."
    }
})

$BtnShare.Add_Click({
    $code = $TxtInvite.Text
    $name = $TxtServerName.Text
    if ($code -and $code -ne "--" -and $code -ne "(pending...)") {
        $msg = "Join my Windrose server '$name'! Connect code: $code (Play > Connect to Server)"
        [System.Windows.Clipboard]::SetText($msg)
        Log "Share message copied to clipboard."
    } else { Log "No invite code available yet." }
})

# Config tab
$CfgMaxPlayers.Add_ValueChanged({
    $TxtMaxPlayersVal.Text = "$([int]$CfgMaxPlayers.Value)"
    $script:MaxPlayers = [int]$CfgMaxPlayers.Value
})

$CfgPasswordEnabled.Add_Checked({   $CfgPassword.IsEnabled = $true  })
$CfgPasswordEnabled.Add_Unchecked({ $CfgPassword.IsEnabled = $false })

$CfgPreset.Add_SelectionChanged({
    $sel = $CfgPreset.SelectedItem
    if ($sel -and $sel.Content -eq "Custom") {
        $PanelCustom.Visibility = "Visible"
    } else {
        $PanelCustom.Visibility = "Collapsed"
    }
})


$sliderPairs = @(
    @{ Slider = $SlMobHealth;  Label = $ValMobHealth  },
    @{ Slider = $SlMobDamage;  Label = $ValMobDamage  },
    @{ Slider = $SlShipHealth; Label = $ValShipHealth },
    @{ Slider = $SlShipDamage; Label = $ValShipDamage },
    @{ Slider = $SlBoarding;   Label = $ValBoarding   },
    @{ Slider = $SlCoopStats;  Label = $ValCoopStats  },
    @{ Slider = $SlCoopShip;   Label = $ValCoopShip   }
)
foreach ($pair in $sliderPairs) {
    $lbl = $pair.Label
    $pair.Slider.Add_ValueChanged({
        $lbl.Text = [Math]::Round($this.Value, 1).ToString("F1")
    }.GetNewClosure())
}

$BtnSaveConfig.Add_Click({
    $p = Get-ServerProcess
    if ($p) { $TxtConfigStatus.Text = "Stop the server before saving config."; $TxtConfigStatus.Foreground = [System.Windows.Media.Brushes]::Tomato; return }
    try {
        # Read existing config to preserve fields we don't edit
        $existingRoot  = $null
        $existingInner = $null
        if (Test-Path $ConfigPath) {
            try {
                $existingRoot  = Get-Content $ConfigPath -Raw | ConvertFrom-Json
                $existingInner = $existingRoot.ServerDescription_Persistent
            } catch {}
        }
        $innerObj = [ordered]@{}
        # Preserve read-only fields
        foreach ($field in @('PersistentServerId','InviteCode','WorldIslandId')) {
            if ($existingInner -and $existingInner.PSObject.Properties[$field]) {
                $innerObj[$field] = $existingInner.$field
            }
        }
        $pw = if ($CfgPasswordEnabled.IsChecked) { $CfgPassword.Text } else { "" }
        $innerObj['IsPasswordProtected'] = ($CfgPasswordEnabled.IsChecked -eq $true)
        $innerObj['Password']            = $pw
        $innerObj['ServerName']          = $CfgName.Text
        $innerObj['MaxPlayerCount']      = [int]$CfgMaxPlayers.Value
        $innerObj['P2pProxyAddress']     = $CfgProxy.Text
        $rootObj = [ordered]@{
            'Version'      = if ($existingRoot -and $existingRoot.PSObject.Properties['Version']) { $existingRoot.Version } else { 1 }
            'DeploymentId' = if ($existingRoot -and $existingRoot.PSObject.Properties['DeploymentId']) { $existingRoot.DeploymentId } else { "" }
            'ServerDescription_Persistent' = $innerObj
        }
        $cfgDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
        if (-not (Test-Path $cfgDir)) { New-Item $cfgDir -ItemType Directory -Force | Out-Null }
        $rootObj | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
        $TxtServerName.Text = $CfgName.Text

        # World config -- write proper nested WorldDescription.json structure
        $wPath = Find-WorldConfig
        if ($wPath) {
            $existingWorld = $null
            if (Test-Path $wPath) {
                try { $existingWorld = Get-Content $wPath -Raw | ConvertFrom-Json } catch {}
            }

            $preset = "Custom"
            $selItem = $CfgPreset.SelectedItem
            if ($selItem) { $preset = $selItem.Content }

            $combatShort = "Normal"
            $cdItem = $CfgCombatDiff.SelectedItem
            if ($cdItem) { $combatShort = $cdItem.Content }

            $floatParams = [ordered]@{
                '{"TagName": "WDS.Parameter.MobHealthMultiplier"}'               = [Math]::Round($SlMobHealth.Value, 2)
                '{"TagName": "WDS.Parameter.MobDamageMultiplier"}'               = [Math]::Round($SlMobDamage.Value, 2)
                '{"TagName": "WDS.Parameter.ShipsHealthMultiplier"}'             = [Math]::Round($SlShipHealth.Value, 2)
                '{"TagName": "WDS.Parameter.ShipsDamageMultiplier"}'             = [Math]::Round($SlShipDamage.Value, 2)
                '{"TagName": "WDS.Parameter.BoardingDifficultyMultiplier"}'      = [Math]::Round($SlBoarding.Value, 2)
                '{"TagName": "WDS.Parameter.Coop.StatsCorrectionModifier"}'     = [Math]::Round($SlCoopStats.Value, 2)
                '{"TagName": "WDS.Parameter.Coop.ShipStatsCorrectionModifier"}' = [Math]::Round($SlCoopShip.Value, 2)
            }
            $boolParams = [ordered]@{
                '{"TagName": "WDS.Parameter.Coop.SharedQuests"}' = ($CfgCoopQuests.IsChecked -eq $true)
                '{"TagName": "WDS.Parameter.EasyExplore"}'       = ($CfgEasyExplore.IsChecked -eq $true)
            }
            $tagParams = [ordered]@{
                '{"TagName": "WDS.Parameter.CombatDifficulty"}' = [ordered]@{
                    'TagName' = "WDS.Parameter.CombatDifficulty.$combatShort"
                }
            }

            # Preserve islandId, WorldName, CreationTime from existing file
            $islandId     = ""
            $worldName    = ""
            $creationTime = 0
            if ($existingWorld -and $existingWorld.WorldDescription) {
                $ewd = $existingWorld.WorldDescription
                if ($ewd.PSObject.Properties['islandId'])     { $islandId     = $ewd.islandId }
                if ($ewd.PSObject.Properties['WorldName'])    { $worldName    = $ewd.WorldName }
                if ($ewd.PSObject.Properties['CreationTime']) { $creationTime = $ewd.CreationTime }
            }

            $worldObj = [ordered]@{
                'Version' = if ($existingWorld -and $existingWorld.PSObject.Properties['Version']) { $existingWorld.Version } else { 1 }
                'WorldDescription' = [ordered]@{
                    'islandId'        = $islandId
                    'WorldName'       = $worldName
                    'CreationTime'    = $creationTime
                    'WorldPresetType' = $preset
                    'WorldSettings'   = [ordered]@{
                        'BoolParameters'  = $boolParams
                        'FloatParameters' = $floatParams
                        'TagParameters'   = $tagParams
                    }
                }
            }
            $worldObj | ConvertTo-Json -Depth 10 | Set-Content $wPath -Encoding UTF8
        }
        $TxtConfigStatus.Text = "Config saved at $(Get-Date -Format 'HH:mm:ss')."
        $TxtConfigStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
    } catch {
        $TxtConfigStatus.Text = "Error: $_"
        $TxtConfigStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
    }
})

$BtnReloadConfig.Add_Click({
    Read-ServerConfig
    Read-WorldConfig
    $TxtConfigStatus.Text = "Config reloaded from disk."
    $TxtConfigStatus.Foreground = $script:BrushGrayText
})

$BtnOpenWorldJson.Add_Click({
    $wPath = Find-WorldConfig
    if ($wPath -and (Test-Path $wPath)) { Start-Process notepad.exe $wPath }
    else { Log "WorldDescription.json not found." }
})

# Log filter buttons
function Set-LogFilter($filterName, $activeBtn) {
    $script:logFilter = $filterName
    foreach ($btn in @($BtnFAll, $BtnFPlayers, $BtnFWarn, $BtnFErrors)) {
        $btn.Background = if ($btn -eq $activeBtn) { $script:BrushBlueBtn } else { $script:BrushGrayBtn }
    }
    Refresh-LogViewer
}
$BtnFAll.Add_Click({     Set-LogFilter "All"     $BtnFAll })
$BtnFPlayers.Add_Click({ Set-LogFilter "Players" $BtnFPlayers })
$BtnFWarn.Add_Click({    Set-LogFilter "Warn"    $BtnFWarn })
$BtnFErrors.Add_Click({  Set-LogFilter "Errors"  $BtnFErrors })

# Console tab
$BtnSendCmd.Add_Click({
    $cmd = $TxtCommand.Text.Trim()
    if (-not $cmd) { return }
    Add-ConsoleEntry $cmd $true
    if (Send-ServerCommand $cmd) {
        $TxtConsoleStatus.Text = ""
    } else {
        $TxtConsoleStatus.Text = "Failed to send command."
    }
    $TxtCommand.Clear()
})

$TxtCommand.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Return) { $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) }
})

$BtnCmdSave.Add_Click({    $TxtCommand.Text = "save world";    $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })
$BtnCmdPlayers.Add_Click({ $TxtCommand.Text = "list players";  $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })
$BtnCmdInfo.Add_Click({    $TxtCommand.Text = "logs";          $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })
$BtnCmdQuit.Add_Click({    $TxtCommand.Text = "quit";          $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })

# Tools tab
$BtnBackup.Add_Click({
    if (-not (Test-Path $SavesBase)) { Log "Saves folder not found: $SavesBase"; return }
    try {
        $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $zipPath = "$BackupDir\Backup_$stamp.zip"
        if (-not (Test-Path $BackupDir)) { New-Item $BackupDir -ItemType Directory -Force | Out-Null }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($SavesBase, $zipPath)
        $script:lastBackupStamp = $stamp
        $TxtLastBackup.Text = "Last backup: $stamp"
        Log "Backup created: $zipPath"
    } catch { Log "Backup error: $_" }
})

$BtnOpenBackups.Add_Click({ Start-Process explorer.exe $BackupDir })

# --- Auto-backup ---
function Get-BackupIntervalHours {
    switch ($CmbBackupInterval.SelectedIndex) {
        0 { 1 }
        1 { 4 }
        2 { 8 }
        3 { 16 }
        4 { 24 }
        default { 4 }
    }
}

function Start-AutoBackupTimer {
    if ($script:autoBackupTimer -ne $null) { $script:autoBackupTimer.Stop(); $script:autoBackupTimer = $null }
    $hours = Get-BackupIntervalHours
    $script:autoBackupTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:autoBackupTimer.Interval = [TimeSpan]::FromHours($hours)
    $script:autoBackupTimer.Add_Tick({
        if (-not (Test-Path $SavesBase)) { Log "Auto-backup skipped: saves folder not found."; return }
        try {
            $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $zipPath = "$BackupDir\Backup_$stamp.zip"
            if (-not (Test-Path $BackupDir)) { New-Item $BackupDir -ItemType Directory -Force | Out-Null }
            [System.IO.Compression.ZipFile]::CreateFromDirectory($SavesBase, $zipPath)
            $script:lastBackupStamp = $stamp
            $TxtLastBackup.Text = "Last backup: $stamp (auto)"
            Log "Auto-backup created: $zipPath"
        } catch { Log "Auto-backup error: $_" }
        # Update next-backup display
        $nextHours = Get-BackupIntervalHours
        $TxtNextBackup.Text = "Next auto-backup: $(Get-Date (Get-Date).AddHours($nextHours) -Format 'h:mm tt')"
    })
    $script:autoBackupTimer.Start()
    $nextTime = (Get-Date).AddHours($hours)
    $TxtNextBackup.Text = "Next auto-backup: $(Get-Date $nextTime -Format 'h:mm tt')"
    Log "Auto-backup enabled: every $hours hour(s)"
}

function Stop-AutoBackupTimer {
    if ($script:autoBackupTimer -ne $null) { $script:autoBackupTimer.Stop(); $script:autoBackupTimer = $null }
    $TxtNextBackup.Text = ""
    Log "Auto-backup disabled."
}

$ChkAutoBackup.Add_Checked({   Start-AutoBackupTimer })
$ChkAutoBackup.Add_Unchecked({ Stop-AutoBackupTimer })
$CmbBackupInterval.Add_SelectionChanged({
    if ($ChkAutoBackup.IsChecked) { Start-AutoBackupTimer }
})

$BtnClearHistory.Add_Click({
    $HistoryList.Items.Clear()
    if (Test-Path $HistoryFile) { Remove-Item $HistoryFile -Force }
    Log "History cleared."
})

# Update tab
$script:latestVersion        = $null
$script:latestContent        = $null
$script:updateDownloadJob    = $null
$script:updateDownloadTimer  = $null

$BtnCheckUpdate.Add_Click({
    $TxtUpdateStatus.Text = "Checking for updates..."
    $TxtUpdateStatus.Foreground = $script:BrushGrayText
    $BtnUpdate.Visibility = "Collapsed"
    $BtnUpdate.IsEnabled = $true   # reset in case a previous download got stuck
    $BtnCheckUpdate.IsEnabled = $false
    # cancel any in-progress download
    if ($script:updateDownloadTimer -ne $null) { $script:updateDownloadTimer.Stop(); $script:updateDownloadTimer = $null }
    if ($script:updateDownloadJob  -ne $null) { Remove-Job $script:updateDownloadJob -Force -ErrorAction SilentlyContinue; $script:updateDownloadJob = $null }
    $script:updateCheckJob = Start-Job -ScriptBlock {
        param($url)
        try {
            $wc = [System.Net.WebClient]::new()
            $wc.Headers.Add("User-Agent", "Windrose-Server-Manager")
            $content = $wc.DownloadString($url)
            $match = [regex]::Match($content, '^\$AppVersion\s*=\s*"([^"]+)"', [System.Text.RegularExpressions.RegexOptions]::Multiline)
            if ($match.Success) {
                return @{ Version = $match.Groups[1].Value; Content = $content; Error = $null }
            } else {
                return @{ Version = $null; Content = $null; Error = "Could not read version from remote file." }
            }
        } catch {
            return @{ Version = $null; Content = $null; Error = $_.ToString() }
        }
    } -ArgumentList $UpdateUrl

    if ($script:updatePollTimer -ne $null) { $script:updatePollTimer.Stop(); $script:updatePollTimer = $null }
    $script:updatePollTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:updatePollTimer.Interval = [TimeSpan]::FromSeconds(1)
    $script:updatePollTimer.Add_Tick({
        $job = $script:updateCheckJob
        if ($job -eq $null -or $job.State -eq 'Running') { return }
        $script:updatePollTimer.Stop()
        $result = Receive-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -ErrorAction SilentlyContinue
        $script:updateCheckJob = $null
        $BtnCheckUpdate.IsEnabled = $true
        if ($result -and $result.Error) {
            $TxtUpdateStatus.Text = "Error: $($result.Error)"
            $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
            return
        }
        if ($result -and $result.Version -ne $null) {
            $script:latestVersion = $result.Version
            $script:latestContent = $result.Content
            if (([System.Version]$result.Version) -gt ([System.Version]$AppVersion)) {
                $TxtUpdateStatus.Text = "Update available! Remote version $($result.Version), you have $AppVersion. Click Update Now to install."
                $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                $BtnUpdate.Visibility = "Visible"
            } else {
                $TxtUpdateStatus.Text = "You are up to date (version $AppVersion)."
                $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
            }
        }
    })
    $script:updatePollTimer.Start()
})

$BtnUpdate.Add_Click({
    if (-not $script:latestVersion) { $TxtUpdateStatus.Text = "Click Check for Updates first."; return }
    $BtnUpdate.IsEnabled = $false
    $TxtUpdateStatus.Text = "Downloading update..."
    $TxtUpdateStatus.Foreground = $script:BrushGrayText
    # Use script-scope so the timer closure can see the job variable
    if ($script:updateDownloadTimer -ne $null) { $script:updateDownloadTimer.Stop(); $script:updateDownloadTimer = $null }
    if ($script:updateDownloadJob  -ne $null) { Remove-Job $script:updateDownloadJob -Force -ErrorAction SilentlyContinue; $script:updateDownloadJob = $null }
    $script:updateDownloadJob = Start-Job -ScriptBlock {
        param($url)
        try {
            $wc = [System.Net.WebClient]::new()
            $wc.Headers.Add("User-Agent","Windrose-Server-Manager")
            $wc.DownloadString($url)   # output directly -- no hashtable wrapper avoids serialisation issues
        } catch {
            "##ERROR##$($_.ToString())"
        }
    } -ArgumentList $UpdateUrl
    $script:updateDownloadTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:updateDownloadTimer.Interval = [TimeSpan]::FromSeconds(1)
    $script:updateDownloadTimer.Add_Tick({
        $job = $script:updateDownloadJob
        if ($job -eq $null -or $job.State -eq 'Running') { return }
        $script:updateDownloadTimer.Stop()
        $script:updateDownloadTimer = $null
        $content = Receive-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -ErrorAction SilentlyContinue
        $script:updateDownloadJob = $null
        $BtnUpdate.IsEnabled = $true
        if (-not $content) {
            $TxtUpdateStatus.Text = "Download returned empty content. Check your connection."
            $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
            return
        }
        if ($content -like "##ERROR##*") {
            $TxtUpdateStatus.Text = "Download failed: $($content -replace '^##ERROR##','')"
            $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
            return
        }
        try {
            $dest = "$ServerDir\Windrose-Server-Manager.ps1"
            [System.IO.File]::WriteAllText($dest, $content, [System.Text.Encoding]::UTF8)
            $BtnUpdate.Visibility = "Collapsed"
            $TxtUpdateStatus.Text = "Updated to version $($script:latestVersion). Close and relaunch to apply."
            $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
            Log "App updated to $($script:latestVersion) -- restart to apply."
        } catch {
            $TxtUpdateStatus.Text = "Failed to write file: $_"
            $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
        }
    })
    $script:updateDownloadTimer.Start()
})

$BtnPatchNotes.Add_Click({
    $w = [System.Windows.Window]::new()
    $w.Title                  = "Patch Notes"
    $w.Width                  = 520
    $w.Height                 = 520
    $w.MinWidth               = 400
    $w.Background             = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x0F,0x19,0x23))
    $w.WindowStartupLocation  = "CenterOwner"
    $w.Owner                  = $Window
    $w.ResizeMode             = "CanResizeWithGrip"

    $scroll = [System.Windows.Controls.ScrollViewer]::new()
    $scroll.VerticalScrollBarVisibility = "Auto"
    $scroll.Padding = [System.Windows.Thickness]::new(14,10,14,10)

    $stack = [System.Windows.Controls.StackPanel]::new()

    foreach ($ver in $PatchNotes.Keys) {
        $card = [System.Windows.Controls.Border]::new()
        $card.Background     = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x11,0x1E,0x2A))
        $card.BorderBrush    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1E,0x33,0x48))
        $card.BorderThickness= [System.Windows.Thickness]::new(1)
        $card.CornerRadius   = [System.Windows.CornerRadius]::new(6)
        $card.Padding        = [System.Windows.Thickness]::new(12,10,12,10)
        $card.Margin         = [System.Windows.Thickness]::new(0,0,0,6)

        $inner = [System.Windows.Controls.StackPanel]::new()

        $hdr = [System.Windows.Controls.TextBlock]::new()
        $hdr.Text       = "Version $ver"
        $hdr.FontSize   = 13
        $hdr.FontWeight = [System.Windows.FontWeights]::Bold
        $hdr.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD4,0xA8,0x43))
        $hdr.Margin     = [System.Windows.Thickness]::new(0,0,0,6)
        $inner.Children.Add($hdr) | Out-Null

        foreach ($line in $PatchNotes[$ver]) {
            $tb = [System.Windows.Controls.TextBlock]::new()
            $tb.Text        = "  - $line"
            $tb.FontSize    = 12
            $tb.Foreground  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xC0,0xCD,0xD8))
            $tb.TextWrapping= "Wrap"
            $tb.Margin      = [System.Windows.Thickness]::new(0,2,0,0)
            $inner.Children.Add($tb) | Out-Null
        }

        $card.Child = $inner
        $stack.Children.Add($card) | Out-Null
    }

    $scroll.Content  = $stack
    $w.Content       = $scroll
    $w.ShowDialog() | Out-Null
})

# Install tab
$BtnDetectSteam.Add_Click({
    $found = Find-SteamWindrose
    if ($found) {
        $TxtSteamSource.Text = $found
        $TxtInstallLog.Text  = "Found: $found"
    } else {
        $TxtInstallLog.Text = "Could not auto-detect Windrose in Steam libraries."
    }
})

$BtnBrowseSource.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = "Select WindowsServer folder (containing WindroseServer.exe)"
    if ($dlg.ShowDialog() -eq "OK") { $TxtSteamSource.Text = $dlg.SelectedPath }
})

$BtnBrowseDest.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = "Select install destination folder"
    if ($dlg.ShowDialog() -eq "OK") { $TxtInstallDest.Text = $dlg.SelectedPath }
})

$BtnInstall.Add_Click({
    $src = $TxtSteamSource.Text.Trim()
    $dst = $TxtInstallDest.Text.Trim()
    if (-not $src -or -not (Test-Path "$src\WindroseServer.exe")) {
        $TxtInstallLog.Text = "ERROR: Source path invalid or WindroseServer.exe not found at:`n$src"
        return
    }
    if (-not (Test-Path $dst)) {
        try { New-Item $dst -ItemType Directory -Force | Out-Null }
        catch { $TxtInstallLog.Text = "ERROR: Could not create destination: $_"; return }
    }
    $BtnInstall.IsEnabled = $false
    $TxtInstallLog.Text = "Installing from:`n$src`nTo:`n$dst`n`nPlease wait..."
    $script:installCopyJob = Start-Job -ScriptBlock {
        param($s, $d)
        robocopy $s $d /E /IS /IT /NP /LOG+:"$d\install.log"
    } -ArgumentList $src, $dst
    if ($script:installTimer -ne $null) { $script:installTimer.Stop(); $script:installTimer = $null }
    $script:installTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:installTimer.Interval = [TimeSpan]::FromSeconds(2)
    $script:installTimer.Add_Tick({
        $job = $script:installCopyJob
        if ($job -eq $null) { $script:installTimer.Stop(); return }
        $installLogPath = "$($TxtInstallDest.Text.Trim())\install.log"
        if ($job.State -eq 'Running') {
            if (Test-Path $installLogPath) {
                try {
                    $tail = Get-Content $installLogPath -Tail 5 -ErrorAction SilentlyContinue
                    if ($tail) { $TxtInstallLog.Text = "Installing...`n" + ($tail -join "`n") }
                } catch {}
            }
        } else {
            $script:installTimer.Stop()
            Receive-Job $job -ErrorAction SilentlyContinue | Out-Null
            Remove-Job  $job -ErrorAction SilentlyContinue
            $script:installCopyJob = $null
            $BtnInstall.IsEnabled = $true
            $dstExe = "$($TxtInstallDest.Text.Trim())\WindroseServer.exe"
            if (Test-Path $dstExe) {
                $DotInstall.Fill = [System.Windows.Media.Brushes]::LimeGreen
                $TxtInstallStatus.Text = "Server installed successfully."
                $TxtInstallStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                $TxtInstallLog.Text += "`n`nInstall complete!"
                if (-not (Test-Path $ConfigPath)) {
                    $cfgDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
                    if (-not (Test-Path $cfgDir)) { New-Item $cfgDir -ItemType Directory -Force | Out-Null }
                    [ordered]@{
                        Version      = 1
                        DeploymentId = ""
                        ServerDescription_Persistent = [ordered]@{
                            PersistentServerId = ""
                            InviteCode         = ""
                            IsPasswordProtected= $false
                            Password           = ""
                            ServerName         = "My Windrose Server"
                            WorldIslandId      = ""
                            MaxPlayerCount     = 10
                            P2pProxyAddress    = "127.0.0.1"
                        }
                    } | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
                    Read-ServerConfig
                    # Pre-fill step 3 with the default name so it's ready to edit
                    $TxtSetupName.Text = "My Windrose Server"
                }
                Update-SetupWizard
            } else {
                $TxtInstallLog.Text += "`n`nWARNING: Install may have failed. WindroseServer.exe not found at destination."
            }
        }
    })
    $script:installTimer.Start()
})

# ---- WIZARD EVENT HANDLERS ----
$BtnCheckReqs.Add_Click({
    $found = Find-SteamWindrose
    if ($found -or (Test-Path $ServerExe)) {
        $TxtSteamSource.Text = if ($found) { $found } else { $TxtSteamSource.Text }
    }
    Update-SetupWizard
})

$SlSetupMaxPlayers.Add_ValueChanged({
    $TxtSetupMaxVal.Text = [int]$SlSetupMaxPlayers.Value
})

$ChkSetupPassword.Add_Checked({   $TxtSetupPassword.IsEnabled = $true  })
$ChkSetupPassword.Add_Unchecked({ $TxtSetupPassword.IsEnabled = $false })

$BtnSaveSetup.Add_Click({
    $name = $TxtSetupName.Text.Trim()
    if (-not $name) { $TxtSetupStatus.Text = "Please enter a server name."; return }
    try {
        $cfgDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
        if (-not (Test-Path $cfgDir)) { New-Item $cfgDir -ItemType Directory -Force | Out-Null }
        $pw = if ($ChkSetupPassword.IsChecked) { $TxtSetupPassword.Text } else { "" }
        $existingInner = $null
        if (Test-Path $ConfigPath) {
            try { $existingInner = (Get-Content $ConfigPath -Raw | ConvertFrom-Json).ServerDescription_Persistent } catch {}
        }
        [ordered]@{
            Version      = 1
            DeploymentId = ""
            ServerDescription_Persistent = [ordered]@{
                PersistentServerId  = if ($existingInner) { $existingInner.PersistentServerId } else { "" }
                InviteCode          = if ($existingInner) { $existingInner.InviteCode }          else { "" }
                IsPasswordProtected = ($ChkSetupPassword.IsChecked -eq $true)
                Password            = $pw
                ServerName          = $name
                WorldIslandId       = if ($existingInner) { $existingInner.WorldIslandId }       else { "" }
                MaxPlayerCount      = [int]$SlSetupMaxPlayers.Value
                P2pProxyAddress     = if ($existingInner) { $existingInner.P2pProxyAddress }     else { "127.0.0.1" }
            }
        } | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
        $script:step3Saved = $true
        Read-ServerConfig
        $TxtSetupStatus.Text = "Saved!"
        $TxtSetupStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
        Update-SetupWizard
    } catch {
        $TxtSetupStatus.Text = "Error: $_"
        $TxtSetupStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
    }
})

$BtnCopyPorts.Add_Click({
    [System.Windows.Clipboard]::SetText("UDP 7777 and UDP 7778")
    $BtnCopyPorts.Content = "Copied!"
    $copyTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $copyTimer.Interval = [TimeSpan]::FromSeconds(2)
    $copyTimer.Add_Tick({ $BtnCopyPorts.Content = "Copy Ports to Clipboard"; $copyTimer.Stop() })
    $copyTimer.Start()
})

$BtnGoToDashboard.Add_Click({ $MainTabs.SelectedIndex = 0 })

# ---- WATCHDOG TIMER ----
$script:watchdogTick = 0
$script:watchdog = [System.Windows.Threading.DispatcherTimer]::new()
$script:watchdog.Interval = [TimeSpan]::FromSeconds(3)
$script:watchdog.Add_Tick({
    $script:watchdogTick++
    $proc       = Get-ServerProcess
    $uiRunning  = ($TxtStatus.Text.Trim() -eq "Running")
    if ($proc -and -not $uiRunning) {
        Set-UIRunning
        if ($script:StartTime -eq $null) { $script:StartTime = [DateTime]::Now }
    } elseif ($proc -and $uiRunning) {
        Update-Stats
        # Full player list rebuild every 30 s to self-correct missed leave events
        if ($script:watchdogTick % 10 -eq 0) { Refresh-PlayerList }
    } elseif (-not $proc -and $uiRunning) {
        Set-UIStopped
        Log "Server process ended unexpectedly."
        $script:ServerProc  = $null
        if ($ChkAutoRestart.IsChecked) {
            Log "Auto-restarting..."
            Add-ConsoleEntry "Server crashed - auto-restarting..."
            & $script:doRestart
        }
    }
    # Schedule check
    if ($ChkSchedule.IsChecked) {
        $nowHm = (Get-Date -Format "HH:mm")
        $target = $TxtScheduleTime.Text.Trim()
        $today  = [DateTime]::Today
        if ($nowHm -eq $target -and $script:lastScheduleDate -ne $today) {
            $script:lastScheduleDate = $today
            Invoke-RestartWithCountdown $script:doRestart
        }
    }
    # Install check
    if (Test-Path $ServerExe) {
        $DotInstall.Fill = [System.Windows.Media.Brushes]::LimeGreen
        $TxtInstallStatus.Text = "Server installed."
        $TxtInstallStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
    }
})

# ---- LOG TAIL TIMER ----
$script:logTailTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:logTailTimer.Interval = [TimeSpan]::FromSeconds(3)
$script:logTailTimer.Add_Tick({ Update-LogViewer })

# ---- INITIAL STATE ----
$TxtInstallDest.Text = $ServerDir
$TxtCurrentVersion.Text = "Current version: $AppVersion"

$BtnRefreshPlayers.Add_Click({
    if (Get-ServerProcess) { Refresh-PlayerList }
    else { $PlayerList.Items.Clear(); $TxtPlayers.Text = "0 / $($script:MaxPlayers)" }
})

# ---- PLAYER LIST CONTEXT MENU ----
$playerMenu     = [System.Windows.Controls.ContextMenu]::new()
$menuKick       = [System.Windows.Controls.MenuItem]::new()
$menuKick.Header = "Kick Player"
$menuBan        = [System.Windows.Controls.MenuItem]::new()
$menuBan.Header  = "Ban Player"
$playerMenu.Items.Add($menuKick) | Out-Null
$playerMenu.Items.Add($menuBan)  | Out-Null
$PlayerList.ContextMenu = $playerMenu

$playerMenu.Add_Opened({
    $hasSelection = ($PlayerList.SelectedItem -ne $null)
    $menuKick.IsEnabled = $hasSelection
    $menuBan.IsEnabled  = $hasSelection
    if ($hasSelection) {
        $menuKick.Header = "Kick $($PlayerList.SelectedItem)"
        $menuBan.Header  = "Ban $($PlayerList.SelectedItem)"
    } else {
        $menuKick.Header = "Kick Player"
        $menuBan.Header  = "Ban Player"
    }
})

$menuKick.Add_Click({
    $playerName = $PlayerList.SelectedItem
    if (-not $playerName) { return }
    if (Send-ServerCommand "kick $playerName") {
        Add-ConsoleEntry "kick $playerName" $true
        Log "Kick sent: $playerName"
    }
})

$menuBan.Add_Click({
    $playerName = $PlayerList.SelectedItem
    if (-not $playerName) { return }
    $confirm = [System.Windows.MessageBox]::Show(
        "Ban player '$playerName'?`nThis will kick them and prevent reconnection.",
        "Confirm Ban", "YesNo", "Warning")
    if ($confirm -ne "Yes") { return }
    if (Send-ServerCommand "ban $playerName") {
        Add-ConsoleEntry "ban $playerName" $true
        Log "Ban sent: $playerName"
    }
})
Read-ServerConfig
Read-WorldConfig
Load-History

# Check last backup
$lastZip = Get-ChildItem "$BackupDir\Backup_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($lastZip) { $TxtLastBackup.Text = "Last backup: $($lastZip.BaseName -replace 'Backup_','')" }

$script:step3Saved = Test-Path $ConfigPath
if (Test-Path $ServerExe) {
    $DotInstall.Fill = [System.Windows.Media.Brushes]::LimeGreen
    $TxtInstallStatus.Text = "Server installed."
    $TxtInstallStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
} else {
    $DotInstall.Fill = $script:BrushRed
    $TxtInstallStatus.Text = "Server not installed - see Install tab"
    $TxtInstallStatus.Foreground = $script:BrushRed
    $MainTabs.SelectedIndex = 5
}
# Auto-detect Steam source and run initial wizard state
$detected = Find-SteamWindrose
if ($detected) { $TxtSteamSource.Text = $detected }
Update-SetupWizard

$existingProc = Get-ServerProcess
if ($existingProc) {
    $script:ServerProc  = $existingProc
    $script:StartTime   = $existingProc.StartTime
    $script:PrevCpuTime = $existingProc.TotalProcessorTime
    $script:PrevCpuCheck = [DateTime]::Now
    Set-UIRunning
    Update-LogViewer
    Log "Attached to running server."
} else {
    Set-UIStopped
}

$script:watchdog.Start()
$script:logTailTimer.Start()

$Window.ShowDialog() | Out-Null
