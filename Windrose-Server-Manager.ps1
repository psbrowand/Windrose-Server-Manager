Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

$AppVersion  = 4
$UpdateUrl   = "https://raw.githubusercontent.com/psbrowand/Windrose-Server-Manager/main/Windrose-Server-Manager.ps1"

$ServerDir   = $PSScriptRoot
$ServerExe   = "$ServerDir\WindroseServer.exe"
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
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
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
          <TextBlock Grid.Row="1" Text="Connected Players" Style="{StaticResource SectionHead}" Margin="0,0,0,4"/>
          <ListBox x:Name="PlayerList" Grid.Row="2" FontSize="13" Margin="0,0,0,8">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Foreground" Value="#C0CDD8"/>
                <Setter Property="Padding" Value="6,3"/>
              </Style>
            </ListBox.ItemContainerStyle>
          </ListBox>
          <CheckBox x:Name="ChkAutoRestart" Grid.Row="3" Content="Auto-restart server if it crashes"
                    Style="{StaticResource DarkCheck}" Margin="0,4,0,0"/>
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
                          TickFrequency="1" IsSnapToTickEnabled="True" Value="4"
                          VerticalAlignment="Center"/>
                  <TextBlock x:Name="TxtMaxPlayersVal" Grid.Column="2" Text="4"
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
            <Button x:Name="BtnCmdInfo" Content="Server Info" Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
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
            <TextBlock Text="Player History" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <StackPanel>
                <ListBox x:Name="HistoryList" Height="150" Margin="0,0,0,8" FontSize="11" FontFamily="Consolas">
                  <ListBox.ItemContainerStyle>
                    <Style TargetType="ListBoxItem">
                      <Setter Property="Padding" Value="4,2"/>
                    </Style>
                  </ListBox.ItemContainerStyle>
                </ListBox>
                <Button x:Name="BtnClearHistory" Content="Clear History" Background="#5A2020" Style="{StaticResource SmallBtn}" HorizontalAlignment="Left"/>
              </StackPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </TabItem>
      <!-- TAB 6: INSTALL -->
      <TabItem Header="Install">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#0F1923">
          <StackPanel Margin="14,10">
            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
              <Ellipse x:Name="DotInstall" Width="14" Height="14" Fill="#CC3333" Margin="0,0,8,0" VerticalAlignment="Center"/>
              <TextBlock x:Name="TxtInstallStatus" Text="Server not installed" Foreground="#CC3333" FontSize="13" VerticalAlignment="Center"/>
            </StackPanel>
            <TextBlock Text="Steam Source" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <StackPanel>
                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox x:Name="TxtSteamSource" Grid.Column="0" Style="{StaticResource DarkInput}"
                           Margin="0,0,4,0"/>
                  <Button x:Name="BtnDetectSteam" Grid.Column="1" Content="Auto-Detect"
                          Background="#1A4A7A" Style="{StaticResource SmallBtn}" Margin="0,0,4,0"/>
                  <Button x:Name="BtnBrowseSource" Grid.Column="2" Content="Browse..."
                          Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
                </Grid>
              </StackPanel>
            </Border>
            <TextBlock Text="Install Destination" Style="{StaticResource SectionHead}"/>
            <Border Background="#111E2A" BorderBrush="#1E3348" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="TxtInstallDest" Grid.Column="0" Style="{StaticResource DarkInput}"
                         Margin="0,0,4,0"/>
                <Button x:Name="BtnBrowseDest" Grid.Column="1" Content="Browse..."
                        Background="#2A3E55" Style="{StaticResource SmallBtn}"/>
              </Grid>
            </Border>
            <TextBlock Text="Install Log" Style="{StaticResource SectionHead}"/>
            <ScrollViewer Height="160" VerticalScrollBarVisibility="Auto" Margin="0,0,0,10">
              <TextBox x:Name="TxtInstallLog" IsReadOnly="True" FontFamily="Consolas" FontSize="10"
                       Background="#0A1218" Foreground="#90A8B8" BorderBrush="#1E3348" BorderThickness="1"
                       TextWrapping="Wrap" VerticalAlignment="Stretch"/>
            </ScrollViewer>
            <Button x:Name="BtnInstall" Content="Install Server" Background="#1A6B3A"
                    Style="{StaticResource BaseBtn}" HorizontalAlignment="Left" Margin="0,0,0,10"/>
            <TextBlock TextWrapping="Wrap" Foreground="#607080" FontSize="11" Margin="0,0,0,10">
              Note: You must own Windrose on Steam (App ID 3041230). The installer copies files
              from your Steam installation. Steam must have downloaded the dedicated server tools.
            </TextBlock>
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
$PlayerList      = Ctrl 'PlayerList'
$ChkAutoRestart  = Ctrl 'ChkAutoRestart'
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
$ChkSchedule     = Ctrl 'ChkSchedule'
$TxtScheduleTime = Ctrl 'TxtScheduleTime'
$TxtCountdown    = Ctrl 'TxtCountdown'
$HistoryList     = Ctrl 'HistoryList'
$BtnClearHistory = Ctrl 'BtnClearHistory'
$TxtCurrentVersion = Ctrl 'TxtCurrentVersion'
$BtnCheckUpdate  = Ctrl 'BtnCheckUpdate'
$BtnUpdate       = Ctrl 'BtnUpdate'
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
$script:MaxPlayers      = 4
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
$script:ServerStdin     = $null
$script:installCopyJob  = $null
$script:installTimer    = $null
$script:updateCheckJob  = $null
$script:updatePollTimer = $null

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
            if ($val -lt 1) { $val = 1 }
            if ($val -gt 10) { $val = 10 }
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
        $preset = "Medium"
        if ($j.PSObject.Properties['Preset']) { $preset = $j.Preset }
        $matchedItem = $null
        foreach ($item in $CfgPreset.Items) {
            if ($item.Content -eq $preset) { $matchedItem = $item; break }
        }
        if ($matchedItem) { $CfgPreset.SelectedItem = $matchedItem }
        if ($preset -eq "Custom") {
            $PanelCustom.Visibility = "Visible"
            $map = @{
                'MobHealthMultiplier'  = $SlMobHealth
                'MobDamageMultiplier'  = $SlMobDamage
                'ShipHealthMultiplier' = $SlShipHealth
                'ShipDamageMultiplier' = $SlShipDamage
                'BoardingMultiplier'   = $SlBoarding
                'CoopStatsMultiplier'  = $SlCoopStats
                'CoopShipMultiplier'   = $SlCoopShip
            }
            foreach ($key in $map.Keys) {
                if ($j.PSObject.Properties[$key]) {
                    $map[$key].Value = [double]$j.$key
                }
            }
            if ($j.PSObject.Properties['CoopQuests'])    { $CfgCoopQuests.IsChecked  = [bool]$j.CoopQuests }
            if ($j.PSObject.Properties['EasyExplore'])   { $CfgEasyExplore.IsChecked = [bool]$j.EasyExplore }
            if ($j.PSObject.Properties['CombatDifficulty']) {
                foreach ($item in $CfgCombatDiff.Items) {
                    if ($item.Content -eq $j.CombatDifficulty) { $CfgCombatDiff.SelectedItem = $item; break }
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

        foreach ($line in $newLines) {
            $script:logBuffer.Add($line)
            $low = $line.ToLower()
            if ($low -match 'lognet: join succeeded:') {
                $playerName = ""
                if ($line -match 'Join succeeded:\s*(.+)') { $playerName = $Matches[1].Trim() }
                if ($playerName) {
                    $script:onlinePlayers.Add($playerName) | Out-Null
                    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] JOINED: $playerName"
                    Add-History $entry
                }
            } elseif ($low -match 'lognet: leave:') {
                $playerName = ""
                if ($line -match 'Leave:\s*(.+)') { $playerName = $Matches[1].Trim() }
                if ($playerName) {
                    $script:onlinePlayers.Remove($playerName) | Out-Null
                    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] LEFT: $playerName"
                    Add-History $entry
                }
            }
        }
        if ($script:logBuffer.Count -gt 1000) {
            $excess = $script:logBuffer.Count - 1000
            $script:logBuffer.RemoveRange(0, $excess)
        }
        Refresh-LogViewer
    } catch {}
}

function Refresh-LogViewer {
    $LogViewer.Items.Clear()
    $filter = $script:logFilter
    foreach ($line in $script:logBuffer) {
        $low = $line.ToLower()
        $include = $false
        switch ($filter) {
            "All"     { $include = $true }
            "Players" { $include = ($low -match 'join succeeded|leave:') }
            "Warn"    { $include = ($low -match 'warning') }
            "Errors"  { $include = ($low -match 'error|fatal') }
        }
        if (-not $include) { continue }
        $tb = [System.Windows.Controls.TextBlock]::new()
        $tb.Text = $line
        $tb.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
        $tb.FontSize = 11
        $tb.TextWrapping = "NoWrap"
        if     ($low -match 'error|fatal')       { $tb.Foreground = [System.Windows.Media.Brushes]::Tomato }
        elseif ($low -match 'warning')            { $tb.Foreground = [System.Windows.Media.Brushes]::Orange }
        elseif ($low -match 'join succeeded')     { $tb.Foreground = [System.Windows.Media.Brushes]::LightGreen }
        elseif ($low -match 'leave:')             { $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xFA,0x80,0x72)) }
        else { $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x70,0x88,0x99)) }
        $LogViewer.Items.Add($tb) | Out-Null
    }
    if ($ChkAutoScroll.IsChecked -and $LogViewer.Items.Count -gt 0) {
        $LogViewer.ScrollIntoView($LogViewer.Items[$LogViewer.Items.Count - 1])
    }
}

function Add-ConsoleEntry($text, $isCommand = $false) {
    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.TextWrapping = "NoWrap"
    $tb.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
    $tb.FontSize = 11
    if ($isCommand) {
        $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD4,0xA8,0x43))
        $tb.Text = "> $text"
    } else {
        $low = $text.ToLower()
        if     ($low -match 'error|fatal')   { $tb.Foreground = [System.Windows.Media.Brushes]::Tomato }
        elseif ($low -match 'warning')       { $tb.Foreground = [System.Windows.Media.Brushes]::Orange }
        elseif ($low -match 'join succeeded'){ $tb.Foreground = [System.Windows.Media.Brushes]::LightGreen }
        elseif ($low -match 'leave:')        { $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xFA,0x80,0x72)) }
        else { $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x70,0x88,0x99)) }
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
            if ($low -match 'lognet: join succeeded:') {
                if ($line -match 'Join succeeded:\s*(.+)') { $online.Add($Matches[1].Trim()) | Out-Null }
            } elseif ($low -match 'lognet: leave:') {
                if ($line -match 'Leave:\s*(.+)') { $online.Remove($Matches[1].Trim()) | Out-Null }
            }
        }
        $reader.Dispose()
        $fs.Dispose()
    } catch {}
    $script:onlinePlayers.Clear()
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

        $PlayerList.Items.Clear()
        foreach ($p in $script:onlinePlayers) { $PlayerList.Items.Add($p) | Out-Null }
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

function Reset-Stats {
    $TxtCpu.Text = "--"
    $TxtRam.Text = "--"
    $TxtPlayers.Text = "--"
    $TxtUptimeBig.Text = "--"
    $TxtUptime.Text = ""
    $PlayerList.Items.Clear()
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
    $DotStatus.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x55,0x55,0x55))
    $TxtStatus.Text = "  Stopped"
    $TxtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x8D,0xA4,0xB5))
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
    $psi.FileName              = $ServerExe
    $psi.Arguments             = "-log"
    $psi.WorkingDirectory      = $ServerDir
    $psi.UseShellExecute       = $false
    $psi.RedirectStandardInput = $true
    $psi.CreateNoWindow        = $false
    $script:ServerProc  = [Diagnostics.Process]::Start($psi)
    $script:ServerStdin = $script:ServerProc.StandardInput
}

$script:doRestart = {
    Stop-AllServerProcesses
    Start-Sleep -Milliseconds 1500
    Start-ServerProcess
    $script:StartTime = [DateTime]::Now
    $script:logPosition = 0L
    $script:logBuffer.Clear()
    $script:onlinePlayers.Clear()
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
        $script:StartTime = [DateTime]::Now
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
    Stop-AllServerProcesses
    $script:ServerStdin = $null
    $script:ServerProc  = $null
    $script:StartTime   = $null
    Set-UIStopped
    $BtnCancelRestart.Visibility = "Collapsed"
    Log "Server stopped."
})

$BtnRestart.Add_Click({
    Invoke-RestartWithCountdown $script:doRestart
})

$BtnSave.Add_Click({
    $p = Get-ServerProcess
    if (-not $p) { Log "Server not running."; return }
    try {
        if ($script:ServerStdin -ne $null) {
            $script:ServerStdin.WriteLine("SaveWorld")
            $script:ServerStdin.Flush()
            Log "Save command sent."
            Add-ConsoleEntry "SaveWorld command sent."
        } else { Log "No stdin handle." }
    } catch { Log "Save error: $_" }
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

        # World config
        $wPath = Find-WorldConfig
        if ($wPath) {
            $preset = "Medium"
            $selItem = $CfgPreset.SelectedItem
            if ($selItem) { $preset = $selItem.Content }
            $worldObj = [ordered]@{ 'Preset' = $preset }
            if ($preset -eq "Custom") {
                $worldObj['MobHealthMultiplier']  = [Math]::Round($SlMobHealth.Value, 2)
                $worldObj['MobDamageMultiplier']  = [Math]::Round($SlMobDamage.Value, 2)
                $worldObj['ShipHealthMultiplier'] = [Math]::Round($SlShipHealth.Value, 2)
                $worldObj['ShipDamageMultiplier'] = [Math]::Round($SlShipDamage.Value, 2)
                $worldObj['BoardingMultiplier']   = [Math]::Round($SlBoarding.Value, 2)
                $worldObj['CoopStatsMultiplier']  = [Math]::Round($SlCoopStats.Value, 2)
                $worldObj['CoopShipMultiplier']   = [Math]::Round($SlCoopShip.Value, 2)
                $worldObj['CoopQuests']           = [bool]$CfgCoopQuests.IsChecked
                $worldObj['EasyExplore']          = [bool]$CfgEasyExplore.IsChecked
                $cdItem = $CfgCombatDiff.SelectedItem
                if ($cdItem) { $worldObj['CombatDifficulty'] = $cdItem.Content } else { $worldObj['CombatDifficulty'] = "Normal" }
            }
            $worldObj | ConvertTo-Json -Depth 5 | Set-Content $wPath -Encoding UTF8
        }
        $TxtConfigStatus.Text = "Config saved at $(Get-Date -Format 'HH:mm:ss')."
        $TxtConfigStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
    } catch {
        $TxtConfigStatus.Text = "Error: $_"
        $TxtConfigStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
    }
})

$BtnOpenWorldJson.Add_Click({
    $wPath = Find-WorldConfig
    if ($wPath -and (Test-Path $wPath)) { Start-Process notepad.exe $wPath }
    else { Log "WorldDescription.json not found." }
})

# Log filter buttons
$BtnFAll.Add_Click({
    $script:logFilter = "All"
    $BtnFAll.Background     = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x4A,0x7A))
    $BtnFPlayers.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFWarn.Background    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFErrors.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    Refresh-LogViewer
})
$BtnFPlayers.Add_Click({
    $script:logFilter = "Players"
    $BtnFAll.Background     = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFPlayers.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x4A,0x7A))
    $BtnFWarn.Background    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFErrors.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    Refresh-LogViewer
})
$BtnFWarn.Add_Click({
    $script:logFilter = "Warn"
    $BtnFAll.Background     = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFPlayers.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFWarn.Background    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x4A,0x7A))
    $BtnFErrors.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    Refresh-LogViewer
})
$BtnFErrors.Add_Click({
    $script:logFilter = "Errors"
    $BtnFAll.Background     = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFPlayers.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFWarn.Background    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2A,0x3E,0x55))
    $BtnFErrors.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1A,0x4A,0x7A))
    Refresh-LogViewer
})

# Console tab
$BtnSendCmd.Add_Click({
    $cmd = $TxtCommand.Text.Trim()
    if (-not $cmd) { return }
    Add-ConsoleEntry $cmd $true
    try {
        if ($script:ServerStdin -ne $null) {
            $script:ServerStdin.WriteLine($cmd)
            $script:ServerStdin.Flush()
            $TxtConsoleStatus.Text = ""
        } else {
            $TxtConsoleStatus.Text = "Server stdin not available."
        }
    } catch {
        $TxtConsoleStatus.Text = "Error: $_"
    }
    $TxtCommand.Clear()
})

$TxtCommand.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Return) { $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) }
})

$BtnCmdSave.Add_Click({    $TxtCommand.Text = "SaveWorld";    $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })
$BtnCmdPlayers.Add_Click({ $TxtCommand.Text = "listplayers";  $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })
$BtnCmdInfo.Add_Click({    $TxtCommand.Text = "stat unit";    $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })
$BtnCmdQuit.Add_Click({    $TxtCommand.Text = "quit";         $BtnSendCmd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })

# Tools tab
$BtnBackup.Add_Click({
    if (-not (Test-Path $SavesBase)) { Log "Saves folder not found."; return }
    try {
        $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $zipPath = "$BackupDir\Backup_$stamp.zip"
        [System.IO.Compression.ZipFile]::CreateFromDirectory($SavesBase, $zipPath)
        $TxtLastBackup.Text = "Last backup: $stamp"
        Log "Backup created: $zipPath"
    } catch { Log "Backup error: $_" }
})

$BtnOpenBackups.Add_Click({ Start-Process explorer.exe $BackupDir })

$BtnClearHistory.Add_Click({
    $HistoryList.Items.Clear()
    if (Test-Path $HistoryFile) { Remove-Item $HistoryFile -Force }
    Log "History cleared."
})

# Update tab
$script:latestVersion   = $null
$script:latestContent   = $null

$BtnCheckUpdate.Add_Click({
    $TxtUpdateStatus.Text = "Checking for updates..."
    $TxtUpdateStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x8D,0xA4,0xB5))
    $BtnUpdate.Visibility = "Collapsed"
    $BtnCheckUpdate.IsEnabled = $false
    $script:updateCheckJob = Start-Job -ScriptBlock {
        param($url)
        try {
            $wc = [System.Net.WebClient]::new()
            $wc.Headers.Add("User-Agent", "Windrose-Server-Manager")
            $content = $wc.DownloadString($url)
            $match = [regex]::Match($content, '^\$AppVersion\s*=\s*(\d+)', [System.Text.RegularExpressions.RegexOptions]::Multiline)
            if ($match.Success) {
                return @{ Version = [int]$match.Groups[1].Value; Content = $content; Error = $null }
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
            if ($result.Version -gt $AppVersion) {
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
    if (-not $script:latestContent) { $TxtUpdateStatus.Text = "No update downloaded yet. Click Check for Updates first."; return }
    try {
        $dest = "$ServerDir\Windrose-Server-Manager.ps1"
        [System.IO.File]::WriteAllText($dest, $script:latestContent, [System.Text.Encoding]::UTF8)
        $BtnUpdate.Visibility = "Collapsed"
        $TxtUpdateStatus.Text = "Update installed (version $($script:latestVersion)). Close and relaunch the app to apply it."
        $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
        Log "App updated to version $($script:latestVersion). Please restart."
    } catch {
        $TxtUpdateStatus.Text = "Failed to write update: $_"
        $TxtUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
    }
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
                            MaxPlayerCount     = 4
                            P2pProxyAddress    = "127.0.0.1"
                        }
                    } | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
                }
            } else {
                $TxtInstallLog.Text += "`n`nWARNING: Install may have failed. WindroseServer.exe not found at destination."
            }
        }
    })
    $script:installTimer.Start()
})

# ---- WATCHDOG TIMER ----
$script:watchdog = [System.Windows.Threading.DispatcherTimer]::new()
$script:watchdog.Interval = [TimeSpan]::FromSeconds(10)
$script:watchdog.Add_Tick({
    $proc       = Get-ServerProcess
    $uiRunning  = ($TxtStatus.Text.Trim() -eq "Running")
    if ($proc -and -not $uiRunning) {
        Set-UIRunning
        if ($script:StartTime -eq $null) { $script:StartTime = [DateTime]::Now }
    } elseif ($proc -and $uiRunning) {
        Update-Stats
    } elseif (-not $proc -and $uiRunning) {
        Set-UIStopped
        Log "Server process ended unexpectedly."
        $script:ServerStdin = $null
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
Read-ServerConfig
Read-WorldConfig
Load-History

# Check last backup
$lastZip = Get-ChildItem "$BackupDir\Backup_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($lastZip) { $TxtLastBackup.Text = "Last backup: $($lastZip.BaseName -replace 'Backup_','')" }

if (Test-Path $ServerExe) {
    $DotInstall.Fill = [System.Windows.Media.Brushes]::LimeGreen
    $TxtInstallStatus.Text = "Server installed."
    $TxtInstallStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
} else {
    $DotInstall.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xCC,0x33,0x33))
    $TxtInstallStatus.Text = "Server not installed - see Install tab"
    $TxtInstallStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xCC,0x33,0x33))
    $MainTabs.SelectedIndex = 5
}

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
