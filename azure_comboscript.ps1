Add-Type -AssemblyName System.Web
Add-Type -AssemblyName PresentationFramework

# Test For Modules
if(-not(Get-Module ExchangeOnlineManagement -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement v2.0.3 - view http://worksmart.link/7l for details')
    Exit
}

if(-not(Get-Module AzureAD -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install AzureAD - view http://worksmart.link/7l for details')
    Exit
}

if(-not(Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install SharePoint Online - view http://worksmart.link/7l for details')
    Exit
}

### Start XAML and Reader to use WPF, as well as declare variables for use
[xml]$xaml = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="Azure Comboscript" Height="600" Width="525" ResizeMode="NoResize" WindowStyle="ThreeDBorderWindow">

    <Grid ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto">
        <Grid.RowDefinitions>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <TabControl Name="Tabs" HorizontalAlignment="Left" Height="565" Margin="10,0,0,0" VerticalAlignment="Top" Width="499">
            <TabItem Name="ResetTab" Header="Reset Password">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Please Pick A User, Then Enter A Password" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Height="25" Width="243"/>
                    <TextBox Name="PasswordTextBox" HorizontalAlignment="Left" Height="25" Margin="10,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="473" TabIndex="1"/>
                    <Button Name="PasswordGoButton" Content="Go" HorizontalAlignment="Left" Margin="10,457,0,0" VerticalAlignment="Top" Width="473" Height="50" IsEnabled="False" TabIndex="0"/>
                    <TextBox Name="UserTextBox" HorizontalAlignment="Left" Height="23" Margin="258,55,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="225" IsReadOnly="True" Background="#FFC8C8C8"/>
                    <Button Name="UserButton" Content="Pick User" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Width="243" Height="54"/>
                    <Label Content="Enter Password Below, Go Button Activates Once Password And User Are Selected" HorizontalAlignment="Left" Margin="10,99,0,0" VerticalAlignment="Top" Width="473"/>
                    <RichTextBox Name="PasswordRichTextBox" HorizontalAlignment="Left" Height="212" Margin="10,240,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                    <CheckBox Name="PasswordResetCheckbox" Content="Force Reset of Password on Login?" HorizontalAlignment="Left" Margin="10,220,0,0" VerticalAlignment="Top" RenderTransformOrigin="-1.221,-1.703" Width="473"/>
                    <Button Name="PasswordReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="258,10,0,0" VerticalAlignment="Top" Width="225" Height="25"/>
                    <Button Name="RandomPasswordButton" Content="Generate Random Password" HorizontalAlignment="Left" Margin="213,191,0,0" VerticalAlignment="Top" Width="270" Height="24"/>
                    <Slider Name="PasswordLengthSlider" HorizontalAlignment="Left" Margin="10,191,0,0" VerticalAlignment="Top" RenderTransformOrigin="-19.075,-4.255" Width="169" Maximum="20" Minimum="12" IsSnapToTickEnabled="True" TickPlacement="BottomRight" Foreground="Cyan"/>
                    <TextBox Name="PasswordLengthTextBox" HorizontalAlignment="Left" Height="23" Margin="184,191,0,0" TextWrapping="Wrap" Text="{Binding ElementName=PasswordLengthSlider, Path=Value, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Top" Width="24" IsReadOnly="True"/>
                    <Label Content="Select Length Of Password For Random Generation If Desired" HorizontalAlignment="Left" Margin="10,160,0,0" VerticalAlignment="Top" Width="473"/>
                </Grid>
            </TabItem>
            <TabItem Name="MailboxTab" Header="Mailboxes">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button Name="MailboxReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <Label Content="Pick a Mailbox to access." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Height="28" Width="473"/>
                    <Button Name="MailboxGoButton" Content="Set Mailbox Permissions For User(s)" HorizontalAlignment="Left" Margin="10,457,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <RichTextBox Name="MailboxRichTextBox" HorizontalAlignment="Center" Height="112" Margin="0,340,0,0" VerticalAlignment="Top" Width="473" Foreground="Cyan" Background="#FF646464" IsReadOnly="True" HorizontalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="MailboxButton" Content="Pick Mailbox" HorizontalAlignment="Left" Margin="10,98,0,0" VerticalAlignment="Top" Width="231" Height="50"/>
                    <TextBox Name="MailboxTextBox" HorizontalAlignment="Left" Height="23" Margin="246,114,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="237" Background="#FFC8C8C8" IsReadOnly="True"/>
                    <Label Content="Checked permissions will be set, unchecked permissions will be removed.  Output will&#xD;&#xA;always show updated, as it overwrites without error." HorizontalAlignment="Center" Margin="0,205,0,0" VerticalAlignment="Top" Width="473"/>
                    <CheckBox Name="FullAccessCheckbox" Content="Full Access" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Center" IsChecked="True"/>
                    <CheckBox Name="SendAsCheckbox" Content="Send As" HorizontalAlignment="Left" Margin="164,0,0,0" VerticalAlignment="Center" IsChecked="True"/>
                    <CheckBox Name="SendOnBehalfCheckbox" Content="Send On Behalf" HorizontalAlignment="Left" Margin="298,0,0,0" VerticalAlignment="Center" RenderTransformOrigin="0.811,0.444" IsChecked="True"/>
                    <GroupBox Header="AutoMap (Full Access Only)" HorizontalAlignment="Left" Height="64" Margin="10,271,0,0" VerticalAlignment="Top" Width="252">
                        <Grid HorizontalAlignment="Left" Height="44" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="232">
                            <RadioButton Name="AutoMapYesRadioButton" Content="Yes" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" IsChecked="True"/>
                            <RadioButton Name="AutoMapNoRadioButton" Content="No" HorizontalAlignment="Left" Margin="61,10,0,0" VerticalAlignment="Top"/>
                            <CheckBox Name="ForceRemapCheckbox" Content="Force Remap?" HorizontalAlignment="Left" Margin="118,10,0,0" VerticalAlignment="Top"/>
                        </Grid>
                    </GroupBox>
                    <Button Name="ConvertToSharedButton" Content="Convert To Shared Mailbox" HorizontalAlignment="Left" Margin="10,155,0,0" VerticalAlignment="Top" Width="231" Height="25" IsEnabled="False"/>
                    <Button Name="ConvertToUserButton" Content="Convert To User Mailbox" HorizontalAlignment="Left" Margin="252,155,0,0" VerticalAlignment="Top" Width="231" Height="25" IsEnabled="False"/>
                    <Button Name="ConvertToEquipmentButton" Content="Convert To Equipment Mailbox" HorizontalAlignment="Left" Margin="10,180,0,0" VerticalAlignment="Top" Width="231" Height="25" IsEnabled="False"/>
                    <Button Name="ConvertToRoomButton" Content="Convert To Room Mailbox" HorizontalAlignment="Left" Margin="252,180,0,0" VerticalAlignment="Top" Width="231" Height="25" IsEnabled="False"/>
                </Grid>
            </TabItem>
            <TabItem Name="GroupTab" Header="Groups">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button Name="GroupReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <Label Content="--When you click Add, you may select multiple users by holding control.  The same&#xD;&#xA;applies to the groups, but it will impact ALL the users you have selected. Please use&#xD;&#xA;caution if you have multiple users with different needed groups.&#xD;&#xA;--When you click Remove, you may select multiple users by holding control.  Groups&#xD;&#xA;work a bit differently, as it will check each user and provide you a prompt to remove&#xD;&#xA;groups based on their membership(s), so you will have to select group(s) for each user." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Height="111" Width="473"/>
                    <Button Name="GroupRemoveButton" Content="Remove User(s) from Group(s)" HorizontalAlignment="Left" Margin="253,457,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                    <Button Name="GroupAddButton" Content="Add User(s) to Group(s)" HorizontalAlignment="Left" Margin="10,457,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                    <RichTextBox Name="GroupRichTextBox" HorizontalAlignment="Left" Height="221" Margin="10,176,0,0" VerticalAlignment="Top" Width="473" IsReadOnly="True" Background="#FF646464" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="ExchangeGroupRemoveButton" Content="Remove User(s) from  Exchange Group(s)" HorizontalAlignment="Left" Margin="253,402,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                    <Button Name="ExchangeGroupAddButton" Content="Add User(s) to Exchange Group(s)" HorizontalAlignment="Left" Margin="10,402,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                </Grid>
            </TabItem>
            <TabItem Name="CalendarTab" Header="Calendars">
                <Grid Background="#FFE5E5E5">
                    <Grid.RowDefinitions>
                        <RowDefinition/>
                    </Grid.RowDefinitions>
                    <Button Name="CalendarReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="20"/>
                    <GroupBox Header="Select Needed Permissions - http://worksmart.link/7f for Permissions Information" HorizontalAlignment="Left" Height="171" Margin="10,148,0,0" VerticalAlignment="Top" Width="473">
                        <Grid HorizontalAlignment="Left" Height="179" Margin="10,10,-2,-41" VerticalAlignment="Top" Width="453">
                            <RadioButton Name="AuthorRadioButton" Content="Author - CI, DOI, EOI, FV, RI" Margin="0,0,237,162"/>
                            <RadioButton Name="ContributorRadioButton" Content="Contributor - CI, FV" Margin="221,0,10,162"/>
                            <RadioButton Name="EditorRadioButton" Content="Editor - CI, DAI, DOI, EAI, EOI, FV, RI" Margin="0,20,237,142"/>
                            <RadioButton Name="NoneRadioButton" Content="None - FV" IsChecked="True" Margin="221,20,10,142"/>
                            <RadioButton Name="NonEditingAuthorRadioButton" Content="NonEditing Author - CI, DOI, FV, RI" Margin="0,40,237,122"/>
                            <RadioButton Name="OwnerRadioButton" Content="Owner - CI, CS, DAI, DOI, EAI, EOI, FC, FO, FV, RI" Margin="0,60,176,102"/>
                            <RadioButton Name="PublishingAuthorRadioButton" Content="Publishing Author - CI, CS, DOI, EOI, FV, RI" Margin="0,80,0,82" HorizontalAlignment="Left" Width="251"/>
                            <RadioButton Name="PublishingEditorRadioButton" Content="Publishing Editor - CI, CS, DAI, DOI, EAI, EOI, FV, RI" Margin="0,100,161,62"/>
                            <RadioButton Name="ReviewerRadioButton" Content="Reviewer - FV, RI" Margin="221,40,10,122"/>
                            <RadioButton Name="AvailabilityOnlyRadioButton" Content="Availability Only" Margin="0,120,0,42"/>
                            <RadioButton Name="LimitedDetailsRadioButton" Content="Limited Details - Availability + Subject and Location" Margin="123,120,10,42"/>
                        </Grid>
                    </GroupBox>
                    <Button Name="CalendarGoButton" Content="Update User's Permissions on Calendar" HorizontalAlignment="Left" Margin="10,457,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <RichTextBox Name="CalendarRichTextBox" HorizontalAlignment="Left" Height="128" Margin="10,324,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                    <Label Content="The permissions below are shorthand, see link for more details.&#xD;&#xA;Create Items, Create Subfolders, Delete All Items, Delete Owned Items, Edit All Items,&#xD;&#xA;Edit Owned Items, Folder Contact, Folder Owner, Folder Visible, Read Items" HorizontalAlignment="Left" Margin="10,86,0,0" VerticalAlignment="Top" Height="57" Width="473"/>
                    <Button Name="CalendarUserButton" Content="Pick User" HorizontalAlignment="Left" Margin="10,35,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button Name="CalendarButton" Content="Pick Calendar" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top" Width="128"/>
                    <TextBox Name="CalendarUserTextBox" HorizontalAlignment="Left" Height="20" Margin="143,35,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="340" Background="#FFC8C8C8" IsReadOnly="True"/>
                    <TextBox Name="CalendarTextBox" HorizontalAlignment="Left" Height="20" Margin="143,60,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="340" Background="#FFC8C8C8" IsReadOnly="True"/>
                    <CheckBox Name="CalendarDefaultCheckbox" Content="Default" HorizontalAlignment="Left" Margin="80,38,0,0" VerticalAlignment="Top"/>
                </Grid>
            </TabItem>
            <TabItem Name="OneDriveTab" Header="OneDrive" Margin="-2,-2,0,0">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Label Content="Click Pick OneDrive to select the OneDrive to be shared.  This may take longer than usual as&#xD;&#xA;it must pull ALL users, not just active ones.  The Share OneDrive will then activate and&#xD;&#xA;allow you to select the user to grant access to the chosen OneDrive folder." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Width="473" Height="60"/>
                    <Button Name="OneDriveGoButton" Content="Share OneDrive" HorizontalAlignment="Left" Margin="10,455,0,0" VerticalAlignment="Top" Width="473" Height="52" IsEnabled="False"/>
                    <RichTextBox Name="OneDriveRichTextBox" HorizontalAlignment="Left" Height="208" Margin="10,242,0,0" VerticalAlignment="Top" Width="473" IsReadOnly="True" Background="#FF646464" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="OneDriveReconnectButton" Content="Reconnect/ChangeTenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <TextBox Name="OneDriveUserTextbox" HorizontalAlignment="Left" Height="23" Margin="258,147,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="225" IsReadOnly="True" Background="#FFC8C8C8"/>
                    <Button Name="OneDriveButton" Content="Pick OneDrive" HorizontalAlignment="Left" Margin="10,130,0,0" VerticalAlignment="Top" Width="243" Height="54"/>
                </Grid>
            </TabItem>
            <TabItem Name="CreateTab" Header="Create User">
                <Grid Background="#FFE5E5E5">
                    <Label Content="First Name" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="FirstNameTextbox" HorizontalAlignment="Left" Height="23" Margin="10,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="0"/>
                    <Label Content="Last Name" HorizontalAlignment="Left" Margin="10,69,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="LastNameTextbox" HorizontalAlignment="Left" Height="23" Margin="10,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="1"/>
                    <Label Content="Username" HorizontalAlignment="Left" Margin="10,128,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="UsernameTextbox" HorizontalAlignment="Left" Height="23" Margin="10,159,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="2"/>
                    <Label Content="@" HorizontalAlignment="Left" Margin="10,187,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="CreatePasswordTextbox" HorizontalAlignment="Left" Height="23" Margin="10,335,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="5"/>
                    <Label Content="Password" HorizontalAlignment="Left" Margin="10,304,0,0" VerticalAlignment="Top" Width="67"/>
                    <ComboBox Name="DomainCombobox" HorizontalAlignment="Left" Margin="10,218,0,0" VerticalAlignment="Top" Width="168" TabIndex="3"/>
                    <Label Content="Usage Location" HorizontalAlignment="Left" Margin="10,247,0,0" VerticalAlignment="Top" Width="91"/>
                    <ComboBox Name="UsageLocationCombobox" HorizontalAlignment="Left" Margin="10,279,0,0" VerticalAlignment="Top" Width="168" TabIndex="4"/>
                    <Label Content="State" HorizontalAlignment="Left" Margin="343,128,0,0" VerticalAlignment="Top" Width="56"/>
                    <TextBox Name="StateTextbox" HorizontalAlignment="Left" Height="23" Margin="343,159,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="9"/>
                    <Label Content="CustomAttribute1" HorizontalAlignment="Left" Margin="343,0,0,0" VerticalAlignment="Center" Width="107"/>
                    <TextBox Name="CustomAttribute1Textbox" HorizontalAlignment="Left" Height="23" Margin="343,276,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="13"/>
                    <RichTextBox Name="CreateRichTextBox" HorizontalAlignment="Left" Height="67" Margin="20,440,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" IsReadOnly="True" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="CreateGoButton" Content="Create User" HorizontalAlignment="Left" Margin="20,394,0,0" VerticalAlignment="Top" Width="463" Height="41" IsEnabled="False" TabIndex="15"/>
                    <Label Content="Country" HorizontalAlignment="Left" Margin="343,187,0,0" VerticalAlignment="Top"/>
                    <TextBox Name="CountryTextbox" HorizontalAlignment="Left" Height="23" Margin="343,217,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="11"/>
                    <Button Name="CreateReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="183,10,0,0" VerticalAlignment="Top" Width="300" Height="26"/>
                    <CheckBox Name="CreateResetPasswordCheckbox" Content="Reset Password on Login?" HorizontalAlignment="Left" Margin="16,363,0,0" VerticalAlignment="Top" TabIndex="6"/>
                    <Slider Name="CreatePasswordLengthSlider" HorizontalAlignment="Center" Margin="0,363,0,0" VerticalAlignment="Top" RenderTransformOrigin="-19.075,-4.255" Width="126" Maximum="20" Minimum="12" IsSnapToTickEnabled="True" TickPlacement="BottomRight" Foreground="Cyan"/>
                    <TextBox Name="CreatePasswordLengthTextBox" HorizontalAlignment="Left" Height="21" Margin="313,363,0,0" TextWrapping="Wrap" Text="{Binding Value, ElementName=CreatePasswordLengthSlider, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Top" Width="23" IsReadOnly="True"/>
                    <Button Name="CreateRandomPasswordButton" Content="Generate Random Password" HorizontalAlignment="Left" Margin="184,335,0,0" VerticalAlignment="Top" Width="154" Height="24"/>
                    <Label Content="CustomAttribute2" HorizontalAlignment="Left" Margin="343,304,0,0" VerticalAlignment="Top" Width="107"/>
                    <TextBox Name="CustomAttribute2Textbox" HorizontalAlignment="Left" Height="23" Margin="343,335,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="14"/>
                    <Label Content="Phone Number" HorizontalAlignment="Left" Margin="189,0,0,0" VerticalAlignment="Center"/>
                    <TextBox Name="PhoneTextbox" HorizontalAlignment="Left" Height="23" Margin="189,276,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="12"/>
                    <Label Content="Zip Code" HorizontalAlignment="Left" Margin="183,187,0,0" VerticalAlignment="Top" Width="57"/>
                    <TextBox Name="ZipTextbox" HorizontalAlignment="Left" Height="23" Margin="189,217,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="10"/>
                    <Label Content="City" HorizontalAlignment="Left" Margin="183,128,0,0" VerticalAlignment="Top" Width="58"/>
                    <TextBox Name="CityTextbox" HorizontalAlignment="Left" Height="23" Margin="189,159,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="8"/>
                    <Label Content="Street Address" HorizontalAlignment="Left" Margin="187,69,0,0" VerticalAlignment="Top" Width="93"/>
                    <TextBox Name="AddressTextbox" HorizontalAlignment="Left" Height="23" Margin="189,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="294" TabIndex="7"/>
                    <Label Content="Create User Button Activates Once Left Side Is Filled In" HorizontalAlignment="Left" Margin="184,36,0,0" VerticalAlignment="Top" Width="299"/>
                </Grid>
            </TabItem>
            <TabItem Name="TerminateTab" Header="Terminate User">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Label Content="Please Select Options Below for User Termination and press Terminate User.  You will&#xD;&#xA;be prompted to select a user and who to share to." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Width="473" Height="43"/>
                    <GroupBox Header="Share OneDrive?" HorizontalAlignment="Left" Height="92" Margin="243,113,0,0" VerticalAlignment="Top" Width="240">
                        <StackPanel HorizontalAlignment="Left" Height="74" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="281">
                            <RadioButton Name="OneDriveNoRadioButton" Content="No" TabIndex="3"/>
                            <RadioButton Name="OneDriveSameRadioButton" Content="To Same User As Shared Mailbox" IsChecked="True" TabIndex="4"/>
                            <RadioButton Name="OneDriveDifferentRadioButton" Content="To Different User As Shared Mailbox" TabIndex="5"/>
                        </StackPanel>
                    </GroupBox>
                    <GroupBox Header="Standard Options" HorizontalAlignment="Left" Height="92" Margin="10,113,0,0" VerticalAlignment="Top" Width="228">
                        <StackPanel HorizontalAlignment="Left" Height="74" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="208">
                            <CheckBox Name="ConvertCheckbox" Content="Convert to Shared Mailbox?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="0"/>
                            <CheckBox Name="RemoveLicensesCheckbox" Content="Remove All Licenses?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="1"/>
                            <CheckBox Name="ShareMailboxCheckbox" Content="Share the Mailbox?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="2"/>
                            <CheckBox Name="ResetMFACheckbox" Content="Reset MFA?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="2"/>
                        </StackPanel>
                    </GroupBox>
                    <Button Name="RemoveGoButton" Content="Terminate User" HorizontalAlignment="Left" Margin="10,455,0,0" VerticalAlignment="Top" Width="473" Height="52"/>
                    <RichTextBox Name="RemoveRichTextBox" HorizontalAlignment="Left" Height="240" Margin="10,210,0,0" VerticalAlignment="Top" Width="473" IsReadOnly="True" Background="#FF646464" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="TerminateReconnectButton" Content="Reconnect/ChangeTenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                </Grid>
            </TabItem>
            <TabItem Name="SpamFilterTab" Header="Spam Filter">
                <Grid Background="#FFE5E5E5">
                    <Label Content="If adding to the Spam Filter, enter a user/domain, and it will add to the correct blacklist.&#xD;&#xA;&#xD;&#xA;If removing, you will need to select the appropriate button, and you will receive a&#xD;&#xA;dialogue box to select and remove." HorizontalAlignment="Center" Margin="0,65,0,0" VerticalAlignment="Top" Height="76" Width="473"/>
                    <Button Name="SpamFilterReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Center" Margin="0,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <RichTextBox Name="SpamFilterRichTextBox" HorizontalAlignment="Center" Height="166" VerticalAlignment="Top" Width="473" Background="#FF646464" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="0,231,0,0">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="SpamFilterRemoveDomainBlacklistButton" Content="Remove Domain(s) from Blacklist" HorizontalAlignment="Left" Margin="253,457,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                    <Button Name="SpamFilterRemoveUserBlacklistButton" Content="Remove User(s) from Blacklist" HorizontalAlignment="Left" Margin="253,402,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                    <Button Name="SpamFilterAddWhitelistButton" Content="Add to Whitelist" HorizontalAlignment="Left" Margin="10,176,0,0" VerticalAlignment="Top" Width="230" Height="50" IsEnabled="False"/>
                    <TextBox Name="SpamFilterTextBox" HorizontalAlignment="Center" Height="25" Margin="0,146,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="473" TabIndex="1"/>
                    <Button Name="SpamFilterAddBlacklistButton" Content="Add to Blacklist" HorizontalAlignment="Left" Margin="253,176,0,0" VerticalAlignment="Top" Width="230" Height="50" IsEnabled="False"/>
                    <Button Name="SpamFilterRemoveDomainWhitelistButton" Content="Remove Domain(s) from Whitelist" HorizontalAlignment="Left" Margin="10,457,0,0" VerticalAlignment="Top" Width="230" Height="50"/>
                    <Button Name="SpamFilterRemoveUserWhitelistButton" Content="Remove User(s) from Whitelist" HorizontalAlignment="Left" Margin="10,402,0,0" VerticalAlignment="Top" Width="230" Height="50"/>

                </Grid>
            </TabItem>
        </TabControl>
    </Grid>

</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
Try{
    $UserForm = [Windows.Markup.XamlReader]::Load($reader)
}
Catch{
    Write-Host "Unable to load Windows.Markup.XamlReader.  Some possible causes for this problem include: .NET Framework is missing, PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."
    Exit
}

#Create Variables For Use In Script Automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $UserForm.FindName($_.Name)}
### End XAML and Variables from XAML

Function Write-RichTextBox {
    Param(
        [System.Windows.Controls.RichTextBox]$TextBox,
        [string]$Text,
        [string]$Color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $TextBox.Document.ContentEnd,$TextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $Text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $Color )
    $TextBox.ScrollToEnd()
}

Function Set-Comboboxes {
    $UsageLocationComboBox.Items.Clear()
    foreach($UsageLocation in $UsageLocations.keys)
    {
        $null = $UsageLocationComboBox.Items.Add($usagelocation)
    }
    $UsageLocationComboBox.SelectedIndex = 0
    
    $domainComboBox.Items.Clear()
    foreach($domain in Get-AzureADDomain){
        $null = $domainComboBox.Items.add($domain.Name)
    }
    $DomainComboBox.SelectedIndex = 0
}

# Friendly Name Lookup Table
$SkuToFriendly = @{
    "9c7bff7a-3715-4da7-88d3-07f57f8d0fb6" = "Dynamics 365 Sales Professional"
    "6070a4c8-34c6-4937-8dfb-39bbc6397a60" = "Meeting Room"
    "8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b" = "Rights Management ADHOC"
    "c42b9cae-ea4f-4ab7-9717-81576235ccac" = "DevPack E5 (No Windows or Audio)"
    "8f0c5670-4e56-4892-b06d-91c085d7004f" = "APP CONNECT IW"
    "0c266dff-15dd-4b49-8397-2bb16070ed52" = "Microsoft 365 Audio Conferencing"
    "2b9c8e7c-319c-43a2-a2a0-48c5c6161de7" = "AZURE ACTIVE DIRECTORY BASIC"
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "AZURE ACTIVE DIRECTORY PREMIUM P1"
    "84a661c4-e949-4bd2-a560-ed7766fcaf2b" = "AZURE ACTIVE DIRECTORY PREMIUM P2"
    "c52ea49f-fe5d-4e95-93ba-1de91d380f89" = "AZURE INFORMATION PROTECTION PLAN 1"
    "295a8eb0-f78d-45c7-8b5b-1eed5ed02dff" = "COMMON AREA PHONE"
    "47794cd0-f0e5-45c5-9033-2eb6b5fc84e0" = "COMMUNICATIONS CREDITS"
    "ea126fc5-a19e-42e2-a731-da9d437bffcf" = "DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION"
    "749742bf-0d37-4158-a120-33567104deeb" = "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
    "cc13a803-544e-4464-b4e4-6d6169a138fa" = "DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
    "8edc2cf8-6438-4fa9-b6e3-aa1660c640cc" = "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
    "1e1a282c-9c54-43a2-9310-98ef728faace" = "DYNAMICS 365 FOR SALES ENTERPRISE EDITION"
    "f2e48cb3-9da0-42cd-8464-4a54ce198ad0" = "DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT"
    "8e7a3d30-d97d-43ab-837c-d7701cef83dc" = "DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION"
    "338148b6-1b11-4102-afb9-f92b6cdc0f8d" = "DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS"
    "b56e7ccc-d5c7-421f-a23b-5c18bdbad7c0" = "DYNAMICS 365 TALENT: ONBOARD"
    "7ac9fe77-66b7-4e5e-9e46-10eed1cff547" = "DYNAMICS 365 TEAM MEMBERS"
    "ccba3cfe-71ef-423a-bd87-b6df3dce59a9" = "DYNAMICS 365 UNF OPS PLAN ENT EDITION"
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "ENTERPRISE MOBILITY + SECURITY E3"
    "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "ENTERPRISE MOBILITY + SECURITY E5"
    "4b9405b0-7788-4568-add1-99614e613b69" = "EXCHANGE ONLINE (PLAN 1)"
    "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "EXCHANGE ONLINE (PLAN 2)"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE"
    "90b5e015-709a-4b8b-b08e-3200f994494c" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
    "7fc0182e-d107-4556-8329-7caaa511197b" = "EXCHANGE ONLINE ESSENTIALS (ExO P1 BASED)"
    "e8f81a67-bd96-4074-b108-cf193eb9433b" = "EXCHANGE ONLINE ESSENTIALS"
    "80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82" = "EXCHANGE ONLINE KIOSK"
    "cb0a98a8-11bc-494c-83d9-c1b1ac65327e" = "EXCHANGE ONLINE POP"
    "061f9ace-7d42-4136-88ac-31dc755f143f" = "INTUNE"
    "b17653a4-2443-4e8c-a550-18249dda78bb" = "Microsoft 365 A1"
    "4b590615-0888-425a-a965-b3bf7789848d" = "MICROSOFT 365 A3 FOR FACULTY"
    "7cfd9a2b-e110-4c39-bf20-c6a3f36a3121" = "MICROSOFT 365 A3 FOR STUDENTS"
    "e97c048c-37a4-45fb-ab50-922fbf07a370" = "MICROSOFT 365 A5 FOR FACULTY"
    "46c119d4-0379-4a9d-85e4-97c66d3f909e" = "MICROSOFT 365 A5 FOR STUDENTS"
    "cdd28e44-67e3-425e-be4c-737fab2899d3" = "MICROSOFT 365 APPS FOR BUSINESS"
    "b214fe43-f5a3-4703-beeb-fa97188220fc" = "MICROSOFT 365 APPS FOR BUSINESS"
    "c2273bd0-dff7-4215-9ef5-2c7bcfb06425" = "MICROSOFT 365 APPS FOR ENTERPRISE"
    "2d3091c7-0712-488b-b3d8-6b97bde6a1f5" = "MICROSOFT 365 AUDIO CONFERENCING FOR GCC"
    "3b555118-da6a-4418-894f-7df1e2096870" = "MICROSOFT 365 BUSINESS BASIC"
    "dab7782a-93b1-4074-8bb1-0e61318bea0b" = "MICROSOFT 365 BUSINESS BASIC"
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "MICROSOFT 365 BUSINESS STANDARD"
    "ac5cef5d-921b-4f97-9ef3-c99076e5470f" = "MICROSOFT 365 BUSINESS STANDARD - PREPAID LEGACY"
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "MICROSOFT 365 BUSINESS PREMIUM"
    "11dee6af-eca8-419f-8061-6864517c1875" = "MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
    "05e9a617-0261-4cee-bb44-138d3ef5d965" = "MICROSOFT 365 E3"
    "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
    "d61d61cc-f992-433f-a577-5bd016037eeb" = "Microsoft 365 E3_USGOV_DOD"
    "ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658" = "Microsoft 365 E3_USGOV_GCCHIGH"
    "184efa21-98c3-4e5d-95ab-d07053a96e67" = "Microsoft 365 E5 Compliance"
    "26124093-3d78-432b-b5dc-48bf992543d5" = "Microsoft 365 E5 Security"
    "44ac31e7-2999-4304-ad94-c948886741d4" = "Microsoft 365 E5 Security for EMS E5"
    "44575883-256e-4a79-9da4-ebe9acabe2b2" = "Microsoft 365 F1"
    "66b55226-6b4f-492c-910c-a3b7a3c9d993" = "Microsoft 365 F3"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "MICROSOFT FLOW FREE"
    "e823ca47-49c4-46b3-b38d-ca11d5abe3d2" = "MICROSOFT 365 G3 GCC"
    "e43b5b99-8dfb-405f-9987-dc307f34bcbd" = "MICROSOFT 365 PHONE SYSTEM"
    "d01d9287-694b-44f3-bcc5-ada78c8d953e" = "MICROSOFT 365 PHONE SYSTEM FOR DOD"
    "d979703c-028d-4de5-acbf-7955566b69b9" = "MICROSOFT 365 PHONE SYSTEM FOR FACULTY"
    "a460366a-ade7-4791-b581-9fbff1bdaa85" = "MICROSOFT 365 PHONE SYSTEM FOR GCC"
    "7035277a-5e49-4abc-a24f-0ec49c501bb5" = "MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH"
    "aa6791d3-bb09-4bc2-afed-c30c3fe26032" = "MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS"
    "1f338bbc-767e-4a1e-a2d4-b73207cc5b93" = "MICROSOFT 365 PHONE SYSTEM FOR STUDENTS"
    "ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61" = "MICROSOFT 365 PHONE SYSTEM FOR TELSTRA"
    "b0e7de67-e503-4934-b729-53d595ba5cd1" = "MICROSOFT 365 PHONE SYSTEM_USGOV_DOD"
    "985fcb26-7b94-475b-b512-89356697be71" = "MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH"
    "440eaaa8-b3e0-484b-a8be-62870b9ba70a" = "MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER"
    "2347355b-4e81-41a4-9c22-55057a399791" = "MICROSOFT 365 SECURITY AND COMPLIANCE FOR FLW"
    "726a0894-2c77-4d65-99da-9775ef05aad1" = "MICROSOFT BUSINESS CENTER"
    "111046dd-295b-4d6d-9724-d52ac90bd1f2" = "MICROSOFT DEFENDER FOR ENDPOINT"
    "906af65a-2970-46d5-9b58-4e9aa50f0657" = "MICROSOFT DYNAMICS CRM ONLINE BASIC"
    "d17b27af-3f49-4822-99f9-56a661538792" = "MICROSOFT DYNAMICS CRM ONLINE"
    "ba9a34de-4489-469d-879c-0f0f145321cd" = "MS IMAGINE ACADEMY"
    "2c21e77a-e0d6-4570-b38a-7ff2dc17d2ca" = "MICROSOFT INTUNE DEVICE FOR GOVERNMENT"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "MICROSOFT POWER APPS PLAN 2 TRIAL"
    "e6025b08-2fa5-4313-bd0a-7e5ffca32958" = "MICROSOFT INTUNE SMB"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "MICROSOFT STREAM"
    "16ddbbfc-09ea-4de2-b1d7-312db6112d70" = "MICROSOFT TEAM (FREE)"
    "710779e8-3d4a-4c88-adb9-386c958d1fdf" = "MICROSOFT TEAMS EXPLORATORY"
    "a4585165-0533-458a-97e3-c400570268c4" = "Office 365 A5 for faculty"
    "ee656612-49fa-43e5-b67e-cb1fdf7699df" = "Office 365 A5 for students"
    "1b1b1f7a-8355-43b6-829f-336cfccb744c" = "Office 365 Advanced Compliance"
    "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
    "18181a46-0d4e-45cd-891e-60aabd171b4e" = "OFFICE 365 E1"
    "6634e0ce-1a9f-428c-a498-f84ec7b8aa2e" = "OFFICE 365 E2"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "OFFICE 365 E3"
    "189a915c-fe4f-4ffa-bde4-85b9628d07a0" = "OFFICE 365 E3 DEVELOPER"
    "b107e5a3-3e60-4c0d-a184-a7e4395eb44c" = "Office 365 E3_USGOV_DOD"
    "aea38a85-9bd5-4981-aa00-616b411205bf" = "Office 365 E3_USGOV_GCCHIGH"
    "1392051d-0cb9-4b7a-88d5-621fee5e8711" = "OFFICE 365 E4"
    "c7df2760-2c81-4ef7-b578-5b5392b571df" = "OFFICE 365 E5"
    "26d45bd9-adf1-46cd-a9e1-51e9a5524128" = "OFFICE 365 E5 WITHOUT AUDIO CONFERENCING"
    "4b585984-651b-448a-9e53-3b10f069cf7f" = "OFFICE 365 F3"
    "535a3a29-c5f0-42fe-8215-d3b9e1f38c4a" = "OFFICE 365 G3 GCC"
    "04a7fb0d-32e0-4241-b4f5-3f7618cd1162" = "OFFICE 365 MIDSIZE BUSINESS"
    "bd09678e-b83c-4d3f-aaba-3dad4abd128b" = "OFFICE 365 SMALL BUSINESS"
    "fc14ec4a-4169-49a4-a51e-2c852931814b" = "OFFICE 365 SMALL BUSINESS PREMIUM"
    "e6778190-713e-4e4f-9119-8b8238de25df" = "ONEDRIVE FOR BUSINESS (PLAN 1)"
    "ed01faf2-1d88-4947-ae91-45ca18703a96" = "ONEDRIVE FOR BUSINESS (PLAN 2)"
    "87bbbc60-4754-4998-8c88-227dca264858" = "POWERAPPS AND LOGIC FLOWS"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "POWER BI (FREE)"
    "45bc2c81-6072-436a-9b0b-3b12eefbc402" = "POWER BI FOR OFFICE 365 ADD-ON"
    "f8a1db68-be16-40ed-86d5-cb42ce701560" = "POWER BI PRO"
    "a10d5e58-74da-4312-95c8-76be4e5b75a0" = "PROJECT FOR OFFICE 365"
    "776df282-9fc0-4862-99e2-70e561b9909e" = "PROJECT ONLINE ESSENTIALS"
    "09015f9f-377f-4538-bbb5-f75ceb09358a" = "PROJECT ONLINE PREMIUM"
    "2db84718-652c-47a7-860c-f10d8abbdae3" = "PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
    "53818b1b-4a27-454b-8896-0dba576410e6" = "PROJECT ONLINE PROFESSIONAL"
    "f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c" = "PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
    "beb6439c-caad-48d3-bf46-0c82871e12be" = "PROJECT PLAN 1"
    "1fc08a02-8b3d-43b9-831e-f76859e04e1a" = "SHAREPOINT ONLINE (PLAN 1)"
    "a9732ec9-17d9-494c-a51c-d6b45b384dcb" = "SHAREPOINT ONLINE (PLAN 2)"
    "b8b749f8-a4ef-4887-9539-c95b1eaa5db7" = "SKYPE FOR BUSINESS ONLINE (PLAN 1)"
    "d42c793f-6c78-4f43-92ca-e8f6a02b035f" = "SKYPE FOR BUSINESS ONLINE (PLAN 2)"
    "d3b4fe1f-9992-4930-8acb-ca6ec609365e" = "SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
    "0dab259f-bf13-4952-b7f8-7db8f131b28d" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
    "54a152dc-90de-4996-93d2-bc47e670fc06" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)"
    "4016f256-b063-4864-816e-d818aad600c9" = "TOPIC EXPERIENCES"
    "de3312e1-c7b0-46e6-a7c3-a515ff90bc86" = "TELSTRA CALLING FOR O365"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd" = "VISIO ONLINE PLAN 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5" = "VISIO ONLINE PLAN 2"
    "4ae99959-6b0f-43b0-b1ce-68146001bdba" = "VISIO PLAN 2 FOR GCC"
    "cb10e6cd-9da4-4992-867b-67546b1db821" = "WINDOWS 10 ENTERPRISE E3"
    "6a0f6da5-0b87-4190-a6ae-9bb5a2b9546a" = "WINDOWS 10 ENTERPRISE E3"
    "488ba24a-39a9-4473-8ee5-19291e71b002" = "Windows 10 Enterprise E5"
    "6470687e-a428-4b7a-bef2-8a291ad947c9" = "WINDOWS STORE FOR BUSINESS"
    "aa2695c9-8d59-4800-9dc8-12e01f1735af" = "NONPROFIT_PORTAL"
}
#Usage Location Lookup Table
$UsageLocations=@{
    "United States" = "US"
    "United Kingdom" = "UK"
}

#Function to Check If Mailbox Exists Before Touching It
function MailboxExistCheck {
    Param(
            [string]$user
        )
    Clear-Variable MailboxExistsCheck -ErrorAction SilentlyContinue
    #Start Mailbox Check Wait Loop
    while ($MailboxExistsCheck -ne $true) {
        try {
            Get-Mailbox $user -ErrorAction Stop
            $MailboxExistsCheck = $true
        }
        catch {
            $null = [System.Windows.MessageBox]::Show("Mailbox Does Not Exist, Waiting 60 Seconds and Trying Again")
            Start-Sleep -Seconds 60
            $MailboxExistsCheck = $false
        }
    }#End Mailbox Check Wait Loop    
}

#Verification Check to enable OK button on Create User Page
function CheckAllBoxes{
    if ( $CreatePasswordTextbox.Text.Length -and ($domainComboBox.SelectedIndex -ge 0) -and $usernameTextbox.Text.Length -and $firstnameTextbox.Text.Length -and $lastnameTextbox.Text.Length )
    {
        $CreateGoButton.IsEnabled = $true
    }
    else {
        $CreateGoButton.IsEnabled = $false
    }
}

### Start Password Tab Functionality
$PasswordTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -ge 2)){
        $PasswordGoButton.IsEnabled = $true
    }
    else{
        $PasswordGoButton.IsEnabled = $false
    }
})

$UserTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -ge 2)){
        $PasswordGoButton.IsEnabled = $true
    }
    else{
        $PasswordGoButton.IsEnabled = $false
    }
})

$PasswordReconnectButton.Add_Click({
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    Try{
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Password Reset
    }
    Connect-AzureAD
    Set-Comboboxes
})

$UserButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
        Set-Comboboxes
    }

    $tempuser = Get-AzureADUser -all $true | Out-GridView -Title "Please Select A User" -Outputmode Single
    $UserTextBox.Text = $tempuser.UserPrincipalName
})

$RandomPasswordButton.Add_Click({
    $PasswordTextBox.Text = [System.Web.Security.Membership]::GeneratePassword($PasswordLengthSlider.Value,($PasswordLengthSlider.Value/2))
})

$PasswordGoButton.Add_Click({
    $securepassword = ConvertTo-SecureString -String $PasswordTextBox.Text -AsPlainText -Force
    Try{
        if($PasswordResetCheckbox.IsChecked -eq $true){
            Set-AzureADUserPassword -ObjectID $UserTextBox.Text -Password $securepassword -ForceChangePasswordNextLogin $true -ErrorAction Stop
        }
        else{
            Set-AzureADUserPassword -ObjectID $UserTextBox.Text -Password $securepassword -ForceChangePasswordNextLogin $false -ErrorAction Stop
        }

        Write-RichtextBox -TextBox $PasswordRichTextBox -Text "SUCCESS:  $($UserTextBox.Text)'s password has been reset to $($PasswordTextBox.Text)`r"
        $UserTextbox.Text = ""
        $PasswordTextBox.Text = ""
    }Catch{
        $message = $_.Exception.Message
        if ($_.Exception.ErrorContent.Message.Value) {
            $message = $_.Exception.ErrorContent.Message.Value
        }
        Write-RichtextBox -TextBox $PasswordRichTextBox -Text "$message`rFAILURE:  Please review above and try again`r" -Color "Red"
    }
})
### End Password Tab Functionality

### Start Mailbox Tab Functionality
$MailboxReconnectButton.Add_Click({
    Try{
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Mailboxes
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$MailboxButton.Add_Click({
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    $TempMailbox = Get-Mailbox -ResultSize Unlimited | Select-Object Displayname,UserPrincipalName,RecipientTypeDetails | Sort-Object Displayname | Out-GridView -Title "Select a Mailbox" -OutputMode Single
    $MailboxTextbox.Text = $TempMailbox.UserPrincipalName
    if($TempMailbox.RecipientTypeDetails -match 'EquipmentMailbox'){
        $ConvertToEquipmentButton.IsEnabled = $false
        $ConvertToRoomButton.IsEnabled = $true
        $ConvertToSharedButton.IsEnabled = $true
        $ConvertToUserButton.IsEnabled = $true
    }
    elseif($TempMailbox.RecipientTypeDetails -match 'RoomMailbox'){
        $ConvertToEquipmentButton.IsEnabled = $true
        $ConvertToRoomButton.IsEnabled = $false
        $ConvertToSharedButton.IsEnabled = $true
        $ConvertToUserButton.IsEnabled = $true
    }
    elseif($TempMailbox.RecipientTypeDetails -match 'SharedMailbox'){        
        $ConvertToEquipmentButton.IsEnabled = $true
        $ConvertToRoomButton.IsEnabled = $true
        $ConvertToSharedButton.IsEnabled = $false
        $ConvertToUserButton.IsEnabled = $true
    }
    elseif($TempMailbox.RecipientTypeDetails -match 'UserMailbox'){
        $ConvertToEquipmentButton.IsEnabled = $true
        $ConvertToRoomButton.IsEnabled = $true
        $ConvertToSharedButton.IsEnabled = $true
        $ConvertToUserButton.IsEnabled = $false
    }
})

$ConvertToEquipmentButton.Add_Click({
    Try{
        Set-Mailbox -Identity $MailboxTextbox.Text -Type Equipment -ErrorAction SilentlyContinue
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) has been converted to an Equipment Mailbox`r" 
        $ConvertToEquipmentButton.IsEnabled = $false
        $ConvertToRoomButton.IsEnabled = $true
        $ConvertToSharedButton.IsEnabled = $true
        $ConvertToUserButton.IsEnabled = $true
    }
    Catch{
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) conversion has failed, please verify manually`r" -Color "Red"
    }
})

$ConvertToRoomButton.Add_Click({
    Try{
        Set-Mailbox -Identity $MailboxTextbox.Text -Type Room -ErrorAction SilentlyContinue
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) has been converted to a Room Mailbox`r" 
        $ConvertToEquipmentButton.IsEnabled = $true
        $ConvertToRoomButton.IsEnabled = $false
        $ConvertToSharedButton.IsEnabled = $true
        $ConvertToUserButton.IsEnabled = $true
    }
    Catch{
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) conversion has failed, please verify manually`r" -Color "Red"
    }
})

$ConvertToSharedButton.Add_Click({
    Try{
        Set-Mailbox -Identity $MailboxTextbox.Text -Type Shared -ErrorAction SilentlyContinue
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) has been converted to a Shared Mailbox`r" 
        $ConvertToEquipmentButton.IsEnabled = $true
        $ConvertToRoomButton.IsEnabled = $true
        $ConvertToSharedButton.IsEnabled = $false
        $ConvertToUserButton.IsEnabled = $true
    }
    Catch{
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) conversion has failed, please verify manually`r" -Color "Red"
    }
})

$ConvertToUserButton.Add_Click({
    Try{
        Set-Mailbox -Identity $MailboxTextbox.Text -Type Regular
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) has been converted to a User Mailbox`r"
        $ConvertToEquipmentButton.IsEnabled = $true
        $ConvertToRoomButton.IsEnabled = $true
        $ConvertToSharedButton.IsEnabled = $true
        $ConvertToUserButton.IsEnabled = $false
    }
    Catch{
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($MailboxTextbox.Text) conversion has failed, please verify manually`r" -Color "Red"
    }
})

$MailboxGoButton.Add_Click({
    Try {
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Clear-Variable SharedMailboxUsers -ErrorAction SilentlyContinue
    $SharedMailboxUsers = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } |  Select-Object Displayname, UserPrincipalName | Sort-Object Displayname | Out-GridView -Title "Select User to Update Access - Hold Ctrl for Multiple" -Passthru
    if($SharedMailboxUsers){
        foreach($SharedMailboxUser in $SharedMailboxUsers){
                if($FullAccessCheckBox.IsChecked){
                    if($AutoMapYesRadioButton.IsChecked){
                        if($ForceRemapCheckBox.IsChecked){
                            Remove-MailboxPermission -Identity $MailboxTextBox.Text -User $SharedMailboxUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -Confirm:$False | Out-Null
                        }
                        Add-MailboxPermission -Identity $MailboxTextBox.Text -User $SharedMailboxUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$False | Out-Null
                    }elseif($AutoMapNoRadioButton.IsChecked){
                        if($ForceRemapCheckBox.IsChecked){
                            Remove-MailboxPermission -Identity $MailboxTextBox.Text -User $SharedMailboxUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -Confirm:$False | Out-Null
                        }
                        Add-MailboxPermission -Identity $MailboxTextBox.Text -User $SharedMailboxUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $false -Confirm:$False | Out-Null
                    }
                    Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($SharedMailboxUser.Displayname) has had Full Access permissions updated on $($MailboxTextBox.Text)`r"
                }else{
                    Remove-MailboxPermission -Identity $MailboxTextBox.Text -User $SharedMailboxUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -Confirm:$False | Out-Null
                    Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($SharedMailboxUser.Displayname) has had Full Access permissions removed on $($MailboxTextBox.Text)`r" -Color "Orange"
                }
                if($SendAsCheckbox.IsChecked){
                    Add-RecipientPermission -Identity $MailboxTextBox.Text -Trustee $SharedMailboxUser.UserPrincipalName -AccessRights SendAs -Confirm:$False
                    Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($SharedMailboxUser.Displayname) has had Send As permissions updated on $($MailboxTextBox.Text)`r"
                }else{
                    Remove-RecipientPermission -Identity $MailboxTextBox.Text -Trustee $SharedMailboxUser.UserPrincipalName -AccessRights SendAs -Confirm:$False
                    Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($SharedMailboxUser.Displayname) has had Send As permissions removed on $($MailboxTextBox.Text)`r" -Color "Orange"
                }
                if($SendOnBehalfCheckbox.IsChecked){
                    Set-Mailbox -Identity $MailboxTextBox.Text -GrantSendOnBehalfTo @{Add="$($SharedMailboxUser.UserPrincipalName)"} -Confirm:$false
                    Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($SharedMailboxUser.Displayname) has had Send On Behalf permissions updated on $($MailboxTextBox.Text)`r"
                }else{
                    Set-Mailbox -Identity $MailboxTextBox.Text -GrantSendOnBehalfTo @{Remove="$($SharedMailboxUser.UserPrincipalName)"} -Confirm:$false
                    Write-RichtextBox -TextBox $MailboxRichTextBox -Text "$($SharedMailboxUser.Displayname) has had Send On Behalf permissions removed on $($MailboxTextBox.Text)`r" -Color "Orange"
                }
        }
        $MailboxTextBox.Text = ""
    }else{
        Write-RichtextBox -TextBox $MailboxRichTextBox -Text "No User Selected" -Color "Red"
    }
})
### End Mailbox Tab Functionality

### Start Group Tab Functionality
$GroupReconnectButton.Add_Click({
    Try{
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Groups
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$ExchangeGroupAddButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-AzureAD        
        Set-Comboboxes
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }

    Clear-Variable Users -ErrorAction SilentlyContinue
    Clear-Variable ExchangeGroups -ErrorAction SilentlyContinue
    
    $Users = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } | Out-GridView -Title "Select User - Hold Ctrl for Multiple" -PassThru
    if($Users){
        $ExchangeGroups = Get-DistributionGroup | Select-Object DisplayName,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups" | Select-Object -Property Displayname,ID
        if ($ExchangeGroups){
            foreach($User in $Users){
                foreach($Group in $Groups){
                    Add-DistributionGroupMember -Identity $ExchangeGroup.ID -Member $user.ObjectID
                    Write-RichtextBox -TextBox $GroupRichTextBox -Text "Added $($user.DisplayName) to $($ExchangeGroup.DisplayName)`r"
                }
            }
        }else{
            Write-RichtextBox -TextBox $GroupRichTextBox -Text "Group Selection Cancelled`r" -Color "Red"
        }
    }else{
        Write-RichtextBox -TextBox $GroupRichTextBox -Text "User Selection Cancelled`r" -Color "Red"
    }    
})

$ExchangeGroupRemoveButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-AzureAD        
        Set-Comboboxes
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }

    Clear-Variable Users -ErrorAction SilentlyContinue
    Clear-Variable ExchangeGroups -ErrorAction SilentlyContinue

    $Users = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } | Out-GridView -Title "Select User - Hold Ctrl for Multiple" -PassThru
    if($Users){
        foreach($User in $Users){
            $ExchangeGroups = Get-DistributionGroup | Select-Object DisplayName,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups" | Select-Object -Property Displayname,ID
            if($ExchangeGroups){
                foreach ($ExchangeGroup in $ExchangeGroups) { 
                    Try
                    {
                        Remove-DistributionGroupMember -ObjectId $ExchangeGroup.Id -MemberId $User.ObjectId
                        Write-RichtextBox -TextBox $GroupRichTextBox -Text "Removed $($user.Displayname) from $($ExchangeGroup.Displayname)`r"
                    }
                    catch
                    {
                        $message = $_.Exception.Message
                        if ($_.Exception.ErrorContent.Message.Value) {
                            $message = $_.Exception.ErrorContent.Message.Value
                        }
                        Write-RichtextBox -TextBox $GroupRichTextBox -Text "Could not remove from group $($ExchangeGroup.Displayname).  Error:  $message)`r" -Color "Red"
                    }
                }
            }else{
                Write-RichtextBox -TextBox $GroupRichTextBox -Text "Group Selection Cancelled" -Color "Red"
            }
        }
    }else{
        Write-RichtextBox -TextBox $GroupRichTextBox -Text "User Selection Cancelled" -Color "Red"
    }
})

$GroupAddButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-AzureAD        
        Set-Comboboxes
    }

    # Pull User ObjectID and Group ObjectID to add member to all groups selected, skipping dynamic
    Clear-Variable Users -ErrorAction SilentlyContinue
    Clear-Variable Groups -ErrorAction SilentlyContinue
    
    $Users = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } | Out-GridView -Title "Select User - Hold Ctrl for Multiple" -PassThru
    if($Users){
        $Groups = Get-AzureADMSGroup -All $true | Where-Object {$_.GroupTypes -ne "DynamicMembership"} | Select-Object DisplayName,Description,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups"
        if ($Groups){
            foreach($User in $Users){
                foreach($Group in $Groups){
                    Add-AzureADGroupMember -ObjectId $Group.Id -RefObjectId $User.ObjectID
                    Write-RichtextBox -TextBox $GroupRichTextBox -Text "$($User.DisplayName) added to $($Group.Displayname)`r"
                }
            }
        }else{
            Write-RichtextBox -TextBox $GroupRichTextBox -Text "Group Selection Cancelled`r" -Color "Red"
        }
    }else{
        Write-RichtextBox -TextBox $GroupRichTextBox -Text "User Selection Cancelled`r" -Color "Red"
    }
})

$GroupRemoveButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-AzureAD
        Set-Comboboxes
    }

    # Pull User ObjectID and Group ObjectID to remove member to all groups selected, skipping dynamic
    Clear-Variable Users -ErrorAction SilentlyContinue
    Clear-Variable Groups -ErrorAction SilentlyContinue
    
    $Users = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } | Out-GridView -Title "Select User - Hold Ctrl for Multiple" -PassThru
    if($Users){
        foreach($User in $Users){
            $Groups = Get-AzureADUserMembership -ObjectId $user.ObjectId | Where-Object {($_.ObjectType -ne "Role") -and ($_.GroupTypes -ne "DynamicMembership")} | Select-Object DisplayName,ObjectId | Sort-Object Displayname | Out-GridView -Title "Select Group - Hold Ctrl for Multiple" -PassThru
            if($Groups){
                foreach ($Group in $Groups) { 
                    Try
                    {
                        Remove-AzureADGroupMember -ObjectId $Group.ObjectId -MemberId $User.ObjectId
                        Write-RichtextBox -TextBox $GroupRichTextBox -Text "Removed $($user.Displayname) from $($Group.Displayname)`r"
                    }
                    catch
                    {
                        $message = $_.Exception.Message
                        if ($_.Exception.ErrorContent.Message.Value) {
                            $message = $_.Exception.ErrorContent.Message.Value
                        }
                        Write-RichtextBox -TextBox $GroupRichTextBox -Text "Could not remove from group $($Group.Displayname).  Error:  $message)`r" -Color "Red"
                    }
                }
            }else{
                Write-RichtextBox -TextBox $GroupRichTextBox -Text "Group Selection Cancelled" -Color "Red"
            }
        }
    }else{
        Write-RichtextBox -TextBox $GroupRichTextBox -Text "User Selection Cancelled" -Color "Red"
    }
})
### End Group Tab Functionality

### Start Calendar Tab Functionality
$CalendarReconnectButton.Add_Click({
    Try{
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Creation
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$CalendarDefaultCheckbox.Add_Checked({
    $CalendarUserButton.IsEnabled = $false
    $CalendarUserTextBox.Text = "Default"
})

$CalendarDefaultCheckbox.Add_Unchecked({
    $CalendarUserButton.IsEnabled = $true
    $CalendarUserTextbox.Text = ""
})

$CalendarUserButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
        Set-Comboboxes
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }

    $TempCalendarUser = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } | Select-Object DisplayName,UserprincipalName | Sort-Object DisplayName | Out-GridView -Title "Select User" -OutputMode Single
    if((Get-Mailbox -Identity $TempCalendarUser.UserPrincipalName | Select-Object RecipientTypeDetails) -match 'SharedMailbox'){
        Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Cannot add Calendar permissions to a Shared Mailbox, please convert to User Mailbox and try again`r" -Color "Red"
    }
    else{
        $CalendarUserTextBox.Text = $TempCalendarUser.UserPrincipalName
    }
})

$CalendarButton.Add_Click({
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    $TempCalendarUser = Get-Mailbox -Filter {(RecipientTypeDetails -eq "SharedMailbox") -or (RecipientTypeDetails -eq "UserMailbox") -or (RecipientTypeDetails -eq "RoomMailbox")} | Select-Object DisplayName,UserPrincipalName | Sort-Object Displayname | Out-GridView -Title "Select Calendar" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
    $TempCalendar = Get-MailboxFolderStatistics $TempCalendarUser | Where-Object {($_.folderpath -match 'Calendar') -and ($_.folderpath -notmatch 'Logging')}| Select-Object Name,Identity,folderpath,foldertype | Out-GridView -Title "Please select a Calendar" -OutputMode Single | Select-Object -ExpandProperty Identity
    $CalendarTextBox.Text = $TempCalendar -replace '^(.*?)\\(.*)','$1:\$2'
})

$CalendarGoButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
        Set-Comboboxes
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    if($AuthorRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Author -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set Author permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Author -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added Author permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"

            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Author Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($ContributorRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Contributor -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set Contributor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Contributor -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added Contributor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Contributor Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($EditorRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Editor -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set Editor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Editor -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added Editor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Editor Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($NoneRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights None -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set None permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights None -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added None permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "None Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($NonEditingAuthorRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights NonEditingAuthor -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set NonEditingAuthor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights NonEditingAuthor -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added NonEditingAuthor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "NonEditingAuthor Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($ReviewerRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Reviewer -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set Reviewer permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Reviewer -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added Reviewer permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Reviewer Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($OwnerRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Owner -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set Owner permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights Owner -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added Owner permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Owner Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($PublishingAuthorRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights PublishingAuthor -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set PublishingAuthor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights PublishingAuthor -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added PublishingAuthor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "PublishingAuthor Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($PublishingEditorRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights PublishingEditor -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set PublishingEditor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights PublishingEditor -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added PublishingEditor permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "PublishingEditor Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($AvailabilityOnlyRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights AvailabilityOnly -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set AvailabilityOnly permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights AvailabilityOnly -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added AvailabilityOnly permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "AvailabilityOnly Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
    elseif($LimitedDetailsRadioButton.IsChecked){
        Try{
            Set-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights LimitedDetails -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully set LimitedDetails permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"
        }Catch{
            Try{
                Add-MailboxFolderPermission -Identity $CalendarTextBox.Text -User $CalendarUserTextBox.Text -AccessRights LimitedDetails -Confirm:$false -ErrorAction Stop
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "Successfully added LimitedDetails permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text)`r"    
            }Catch{
                Write-RichtextBox -TextBox $CalendarRichTextBox -Text "LimitedDetails Permissions for $($CalendarUserTextbox.Text) on $($CalendarTextBox.Text) Not Set`r" -Color "Red"
            }
        }
    }
})
### End Calendar Tab Functionality

### Start OneDrive Tab Functionality
$OneDriveReconnectButton.Add_Click({
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    Connect-AzureAD
    $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
    $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
    Connect-SPOService -Url $AdminSiteURL
    Set-Comboboxes
})

$OneDriveButton.Add_Click({
    Clear-Variable tempuser -ErrorAction SilentlyContinue
    $tempuser = Get-AzureADUser -all $true | Out-GridView -Title "Please Select A User" -Outputmode Single
    $OneDriveUserTextBox.Text = $tempuser.UserPrincipalName
})

$OneDriveUserTextbox.Add_TextChanged({
    if ($OneDriveUserTextbox.Text.Length -ge 2){
        $OneDriveGoButton.IsEnabled = $true
    }
    else{
        $OneDriveGoButton.IsEnabled = $false
    }
})

$OneDriveGoButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
        Set-Comboboxes
    }
    Try{
        Get-SPOHomeSite -ErrorAction Stop | Out-Null
    }
    Catch{
        $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
        $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
        Connect-SPOService -Url $AdminSiteURL
    }

    Clear-Variable SharedOneDriveUser -Erroraction SilentlyContinue
    $SharedOneDriveUser = Get-AzureADUser -All $true | Sort-Object Displayname | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user to share the OneDrive with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
    #Pull Object ID Needed For User Receiving Access To OneDrive And OneDriveSiteURL Dynamically
    if($SharedOneDriveUser){
        Try{
            $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($OneDriveUserTextbox.Text)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url -ErrorAction Stop            
            #Add User Receiving Access To Terminated User's OneDrive
            Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedOneDriveUser -IsSiteCollectionAdmin $True
            Write-RichtextBox -TextBox $OneDriveRichTextBox -Text "$($OneDriveUserTextBox.Text)'s OneDrive Data Shared with $SharedOneDriveUser successfully, link to copy and provide to trustee is $OneDriveSiteURL`r"
            $OneDriveUserTextbox.Text = ""
        }Catch{
            Write-RichtextBox -TextBox $OneDriveRichTextBox -Text "OneDrive URL does not exist, verify user in 365 manually`r" -Color "Red"
        }
    }else{
        Write-RichtextBox -TextBox $OneDriveRichTextBox -Text  "OneDrive Share Cancelled" -Color "Red"
    }
})
### End OneDrive Tab Functionality

### Start User Creation Tab Functionality
$firstnameTextbox.Add_TextChanged({
    CheckAllBoxes
})

$lastnameTextbox.Add_TextChanged({
    CheckAllBoxes
})

$usernameTextbox.Add_TextChanged({
    CheckAllBoxes
})

$CreatePasswordTextbox.Add_TextChanged({
    CheckAllBoxes
})

$DomainCombobox.Add_SelectionChanged({
    CheckAllBoxes
})

$UsageLocationCombobox.Add_SelectionChanged({
    CheckAllBoxes
})

$CreateReconnectButton.Add_Click({
    Try{
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Creation
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$CreateRandomPasswordButton.Add_Click({
    $CreatePasswordTextBox.Text = [System.Web.Security.Membership]::GeneratePassword($CreatePasswordLengthSlider.Value,($CreatePasswordLengthSlider.Value/2))
})

$CreateGoButton.Add_Click({
    Clear-Variable AvailableLicenseCheck -ErrorAction SilentlyContinue
    Clear-Variable UserExists -ErrorAction SilentlyContinue
    Clear-Variable Licenses -ErrorAction SilentlyContinue
    Clear-Variable SelectedLicenses -ErrorAction SilentlyContinue
    Clear-Variable Groups -ErrorAction SilentlyContinue
    Try{
        Get-AzureADUser -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-AzureAD
        Set-Comboboxes
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }
    Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    $Licenses =  Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits
    foreach($License in $Licenses){
        $TempSkuCheck = $skuToFriendly.Item("$($License.SkuID)")
        if($TempSkuCheck){
            $License.SkuPartNumber = $skuToFriendly.Item("$($License.SkuID)")
        }
        else{
            $null = [System.Windows.MessageBox]::Show("Please Submit a Github Issue for Non-Matching SkuPartNumber $($License.SkuID) - $($License.SkuPartNumber): https://github.com/mrobinson-ws/azure_comboscript/issues")
        }
    }
    while($AvailableLicenseCheck -ne $true){
        $SelectedLicenses = $Licenses | Sort-Object SkuPartNumber | Out-GridView -Passthru -Title "Hold Ctrl For Multiple Licenses"
        foreach($SelectedLicense in $SelectedLicenses){
            if($SelectedLicense.Enabled-$SelectedLicense.ConsumedUnits -ge 1){
                $Available = $SelectedLicense.Enabled-$SelectedLicense.ConsumedUnits
                $null = [System.Windows.MessageBox]::Show("You have $Available available $($SelectedLicense.SkuPartNumber) licenses")
                $AvailableLicenseCheck = $true
            }
            elseif($SelectedLicense.Enabled-$SelectedLicense.ConsumedUnits -le 0){
                $null = [System.Windows.MessageBox]::Show("You do not have any $($SelectedLicense.SkuPartNumber) licenses to assign, please acquire licenses and try again","License Check","OKCancel","Warning")
                $AvailableLicenseCheck = $false
            }
        }
        if(-not($SelectedLicenses)){
            Write-RichtextBox -TextBox $CreateRichTextBox -Text "No License Selected`r" -Color "Red"
            Break
        }
    }

    if ($AvailableLicenseCheck -eq $true) {
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $CreatePasswordTextbox.text
        $UPN = $usernameTextbox.Text + "@" + $domainCombobox.Text
        $displayname = $firstnameTextbox.text + " " + $lastnameTextbox.Text
        $usageloc = $UsageLocations[$UsageLocationComboBox.Text]
               
        #Test if User with matching UPN already exists
        try {
            Write-RichtextBox -TextBox $CreateRichTextBox -Text "Testing If Username Does Not Exist`r"
            Get-AzureAdUSer -ObjectID $UPN -ErrorAction Stop | Out-Null
            $UserExists = $true
        }
        #Otherwise, creates user and assigns licenses selected in first step
        catch {
            Write-RichtextBox -TextBox $CreateRichTextBox -Text "Username Does Not Exist, Creating User and Assigning Licenses`r"
            $UserExists = $false
        }

        if($UserExists -ne $true){
            New-AzureADUser -DisplayName $displayname -PasswordProfile $PasswordProfile -UserPrincipalName $UPN -AccountEnabled $true -MailNickName "$($usernameTextbox.Text)" -UsageLocation $usageloc
            foreach($SelectedLicense in $SelectedLicenses){
                $AssignedLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                $AssignedLicense.SkuID = $SelectedLicense.SkuID
                $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses 
                $Licenses.AddLicenses = $AssignedLicense
                Set-AzureADUserLicense -ObjectID $UPN -AssignedLicenses $Licenses
            }
            if([string]::IsNullOrwhiteSpace($AddressTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectId $UPN -StreetAddress $AddressTextbox.Text
            }            
            if([string]::IsNullOrwhiteSpace($cityTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectId $UPN -City $cityTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($stateTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectID $UPN -State $stateTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($ZipTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectID $UPN -PostalCode $ZipTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($countryTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectID $UPN -Country $countryTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($PhoneTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectID $UPN -TelephoneNumber $PhoneTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($CustomAttribute1Textbox.Text) -eq $false){
                MailboxExistCheck($UPN)
                Set-Mailbox $UPN -CustomAttribute1 $CustomAttribute1Textbox.Text
            }
            if([string]::IsNullOrWhiteSpace($CustomAttribute2Textbox.Text) -eq $false){
                MailboxExistCheck($UPN)
                Set-Mailbox $UPN -CustomAttribute2 $CustomAttribute2Textbox.Text
            }
        
            $user = Get-AzureADUser -ObjectID $UPN
            MailboxExistCheck($UPN)
            $Groups = Get-AzureADMSGroup -All $true | Where-Object {$_.GroupTypes -ne "DynamicMembership"} | Select-Object DisplayName,Description,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups" | Select-Object -Property Displayname,Id
            if ($Groups){
                foreach($group in $Groups){
                    Add-AzureADGroupMember -ObjectId $group.Id -RefObjectId $user.ObjectID
                    Write-RichtextBox -TextBox $CreateRichTextBox -Text "Added $($user.DisplayName) to $($group.DisplayName)`r"
                }
                
            }
            else {
                Write-RichtextBox -TextBox $CreateRichTextBox -Text "No Groups Selected`r" -Color "Yellow"
            }

            $ExchangeGroups = Get-DistributionGroup | Select-Object DisplayName,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups" | Select-Object -Property Displayname,ID
            if ($ExchangeGroups){
                foreach($ExchangeGroup in $ExchangeGroups){
                    Add-DistributionGroupMember -Identity $ExchangeGroup.ID -Member $user.ObjectID
                    Write-RichtextBox -TextBox $CreateRichTextBox -Text "Added $($user.DisplayName) to $($ExchangeGroup.DisplayName)`r"
                }
                
            }
            else {
                Write-RichtextBox -TextBox $CreateRichTextBox -Text "No Groups Selected`r" -Color "Yellow"
            }


            Write-RichtextBox -TextBox $CreateRichTextBox -Text "User created for $($firstnameTextbox.Text) $($lastnameTextbox.Text)`r"
            $CustomAttribute2Textbox.Text = ""
            $CustomAttribute1Textbox.Text = ""
            $cityTextbox.Text = ""
            $stateTextbox.Text = ""
            $zipTextbox.Text = ""
            $AddressTextbox.Text = ""
            $countryTextBox.Text = ""
            $firstnameTextbox.Text = ""
            $lastnameTextbox.Text = ""
            $usernameTextbox.Text = ""
            $CreatePasswordTextbox.Text = ""
        }
        Else{
            $ExistingUser = Get-AzureADUser -ObjectID $UPN
            Write-RichtextBox -TextBox $CreateRichTextBox -Text "Username $UPN Exists as $($ExistingUser.DisplayName), Please Review and Try Again`r" -Color "Red"
        }
    }
})
### End User Creation Tab Functionality

### Start User Termination Tab Functionality
$TerminateReconnectButton.Add_Click({
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
    $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
    Connect-SPOService -Url $AdminSiteURL
    Set-Comboboxes
})

$RemoveGoButton.Add_Click({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
        Set-Comboboxes
    }
    Try{
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    #Pull All Azure AD Users and Store In Hash Table Instead Of Calling Get-AzureADUser Multiple Times
    $allUsers = @{}    
    foreach ($user in Get-AzureADUser -All $true){ $allUsers[$user.UserPrincipalName] = $user }
    
    #Request Username(s) To Be Terminated From Script Runner (Hold Ctrl To Select Multiples)
    $username = $allUsers.Values | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-Gridview -OutputMode Single -Title "Please select the user to be terminated" | Select-Object -ExpandProperty UserPrincipalName
    
    ##### Start User(s) Loop #####
    if($username){
        $UserInfo = $allusers[$username]

        #Block Sign In Of User/Force Sign Out Within 60 Minutes
        Set-AzureADUser -ObjectID $UserInfo.ObjectId -AccountEnabled $false
        Write-RichtextBox -TextBox $RemoveRichTextBox -Text "Sign in Blocked for $($UserInfo.ObjectID)`r"

        #Remove All Group Memberships
        $memberships = Get-AzureADUserMembership -ObjectId $username | Where-Object {$_.ObjectType -ne "Role"}| Select-Object DisplayName,ObjectId
        foreach ($membership in $memberships) { 
                $group = Get-AzureADMSGroup -ID $membership.ObjectId
                if ($group.GroupTypes -contains 'DynamicMembership') {
                    Write-RichtextBox -TextBox $RemoveRichTextBox -Text "Skipped $($group.Displayname) as it is dynamic`r" -Color "Yellow"
                }
                else{
                    Try{
                        Remove-AzureADGroupMember -ObjectId $membership.ObjectId -MemberId $UserInfo.ObjectId -ErrorAction Stop
                    }Catch{
                        $message = $_.Exception.Message
                        if ($_.Exception.ErrorContent.Message.Value) {
                            $message = $_.Exception.ErrorContent.Message.Value
                        }
                        Write-RichtextBox -TextBox $RemoveRichTextBox -Text "Could not remove from group $($group.name).  Error:  $message`r" -Color "Yellow"
                    }
             
                }
            }
            Write-RichtextBox -TextBox $RemoveRichTextBox -Text "All non-dynamic groups removed, please check your Downloads folder for the file, it will also open automatically at end of user termination`r"

        #Convert To Shared Mailbox And Hide From GAL When Convert Is Selected, Must Be Done Before Removing Licenses
        if ($ConvertCheckBox.IsChecked -eq $true) {
            Set-Mailbox $username -Type Shared -HiddenFromAddressListsEnabled $true
            Write-RichtextBox -TextBox $RemoveRichTextBox -Text "Mailbox for $username converted to Shared, address hidden from GAL`r"
        }

        #Grant Access To Shared Mailbox When Grant CheckBox Is Selected
        if ($ShareMailboxCheckBox.IsChecked -eq $true) {
            $sharedMailboxUser = $allUsers.Values | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the $username Shared Mailbox with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
            if($sharedMailboxUser){
               Add-MailboxPermission -Identity $username -User $SharedMailboxUser -AccessRights FullAccess -InheritanceType All
                Add-RecipientPermission -Identity $username -Trustee $SharedMailboxUser -AccessRights SendAs -Confirm:$False
                Write-RichtextBox -TextBox $RemoveRichTextBox -Text "Access granted to the $username Shared Mailbox to $sharedMailboxUser`r"
            }
            else{
                Write-RichtextBox -TextBox $RemoveRichTextBox -Text "Cancelled Sharing of Mailbox`r" -Color "Red"
            }
        }

        #Remove All Licenses When Remove Licenses Is Selected
        if ($RemoveLicensesCheckBox.IsChecked -eq $true) {
            $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            if($UserInfo.assignedlicenses){
                $licenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $licenses
            }
            Write-RichtextBox -TextBox $RemoveRichTextBox -Text "All licenses have been removed`r"
        }

        #Test And Connect To Sharepoint Online If Needed
        if ($OneDriveNoRadioButton.IsChecked -ne $true) {
            $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
            $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
            Try{
                Get-SPOSite -ErrorAction Stop | Out-Null
            }Catch{
                Try{
                    Connect-SPOService -Url $AdminSiteURL
                }Catch{
                    Write-RichtextBox -TextBox $RemoveRichTextBox -Text "SharePointOnline Could Not Connect" -Color "Red"
                    Break
                }
            }
        }

        #Share OneDrive With Same User as Shared Mailbox
        if ($OneDriveSameRadioButton.IsChecked) {
            Clear-Variable OneDriveSiteURL -ErrorAction SilentlyContinue
            #Pull OneDriveSiteURL Dynamically And Grant Access
            Try{
                $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url -ErrorAction SilentlyContinue
                #Add User Receiving Access To Terminated User's OneDrive
                Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedMailboxUser -IsSiteCollectionAdmin $True
                Write-RichtextBox -TextBox $RemoveRichTextBox -Text "OneDrive Data Shared with $SharedMailboxUser successfully, link to copy and give to Manager is $OneDriveSiteURL`r"
            }
            Catch{
                Write-RichtextBox -TextBox $RemoveRichTextBox -Text "OneDrive URL does not exist, verify user in 365 manually (OneDrive Tab once confirmed)`r" -Color "Red"
            }
        }
        #Share OneDrive With Different User(s) than Shared Mailbox
        elseif ($OneDriveDifferentRadioButton.IsChecked) {
            Clear-Variable SharedOneDriveUser -Erroraction SilentlyContinue
            $SharedOneDriveUser = $allusers.Values | Sort-Object Displayname | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user to share the OneDrive with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
            #Pull Object ID Needed For User Receiving Access To OneDrive And OneDriveSiteURL Dynamically
            if($SharedOneDriveUser){
                Clear-Variable SharedOneDriveUser -Erroraction SilentlyContinue
                Try{
                    $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url -ErrorAction SilentlyContinue
                    #Add User Receiving Access To Terminated User's OneDrive
                    Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedOneDriveUser -IsSiteCollectionAdmin $True
                    Write-RichtextBox -TextBox $RemoveRichTextBox -Text "OneDrive Data Shared with $SharedOneDriveUser successfully, link to copy and provide to trustee is $OneDriveSiteURL`r"
                }
                Catch{
                    Write-RichtextBox -TextBox $RemoveRichTextBox -Text "OneDrive URL does not exist, verify user in 365 manually (OneDrive Tab once confirmed)`r" -Color "Red"
                }
            }else{
                Write-RichtextBox -TextBox $RemoveRichTextBox -Text  "OneDrive Share Cancelled`r" -Color "Red"
            }
        }

        #Export Groups Removed and OneDrive URL to CSV
        [pscustomobject]@{
            GroupsRemoved    = $memberships.DisplayName -join ','
            OneDriveSiteURL = $OneDriveSiteURL
        } | Export-Csv -Path c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv -NoTypeInformation

        #Open Created CSV File At End Of Loop For Ease Of Copying OneDrive URL To Give
        Start-Process c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv
    }
})

$ConvertCheckbox.Add_Checked({
    $ShareMailboxCheckBox.IsEnabled = $true
})

$ConvertCheckbox.Add_Unchecked({
    $ShareMailboxCheckBox.IsChecked = $false
    $ShareMailboxCheckBox.IsEnabled = $false
})

$ShareMailboxCheckBox.Add_Checked({ 
    $OneDriveSameRadioButton.IsEnabled = $true
})

$ShareMailboxCheckBox.Add_Unchecked({
    $OneDriveSameRadioButton.IsEnabled = $false
    $OneDriveNoRadioButton.IsChecked = $true
})
### End User Termination Tab Functionality

### Start Spam Filter Tab Functionality
$SpamFilterReconnectButton.Add_Click({

})

$SpamFilterTextBox.Add_TextChanged({
    if($SpamFilterTextbox.Text.Length -ge 2){
        $SpamFilterAddWhitelistButton.IsEnabled = $true
        $SpamFilterAddBlacklistButton.IsEnabled = $true
    }else{
        $SpamFilterAddWhitelistButton.IsEnabled = $false
        $SpamFilterAddBlacklistButton.IsEnabled = $false
    }

})

$SpamFilterAddWhitelistButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline
    }
    if([string]::IsNullOrwhiteSpace($SpamFilterTextBox.Text) -eq $false){
        Try{
            Set-HostedContentFilterPolicy Default -AllowedSenderDomains @{Add="$($SpamFilterTextbox.Text)"} -ErrorAction Stop
            Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text  "Added $($SpamFilterTextbox.Text) to Domain Whitelist`r"
            $SpamFilterTextBox.Text = ""
        }
        Catch{
            Try{
                Set-HostedContentFilterPolicy Default -AllowedSenders @{Add="$($SpamFilterTextbox.Text)"} -ErrorAction Stop
                Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text  "Added $($SpamFilterTextbox.Text) to Sender Whitelist`r"
                $SpamFilterTextBox.Text = ""
            }
            Catch{
                Write-RichtextBox -TextBox $SpamFilterTextBox -Text  "$_.Message indicates $($SpamFilterTextbox.Text) is potentially invalid, please review entry and try again`r" -Color "Red"
            }
        }
    }
})

$SpamFilterAddBlacklistButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline
    }
    if([string]::IsNullOrwhiteSpace($SpamFilterTextBox.Text) -eq $false){
        Try{
            Set-HostedContentFilterPolicy Default -BlockedSenderDomains @{Add="$($SpamFilterTextbox.Text)"} -ErrorAction Stop
            Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text  "Added $($SpamFilterTextbox.Text) to Domain Blacklist`r"
            $SpamFilterTextBox.Text = ""
        }
        Catch{
            Try{
                Set-HostedContentFilterPolicy Default -BlockedSenders @{Add="$($SpamFilterTextbox.Text)"} -ErrorAction Stop
                Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text  "Added $($SpamFilterTextbox.Text) to Sender Blacklist`r"
                $SpamFilterTextBox.Text = ""
            }
            Catch{
                Write-RichtextBox -TextBox $SpamFilterTextBox -Text  "$_.Message indicates $($SpamFilterTextbox.Text) is potentially invalid, please review entry and try again`r" -Color "Red"
            }
        }
    }
})

$SpamFilterRemoveUserWhitelistButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline
    }
    Clear-Variable rfwl -ErrorAction SilentlyContinue
    Clear-Variable rfwls -ErrorAction SilentlyContinue
    Try{
        $rfwls = Get-HostedContentFilterPolicy Default
        $rfwls = $rfwls.AllowedSenders | Select-Object -Property Sender | Out-GridView -Passthru -Title "Select Multiple Senders By Holding Ctrl"
        foreach($rfwl in $rfwls){
            Set-HostedContentFilterPolicy Default -AllowedSenders @{Remove="$($rfwl.sender)"}
            Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text "Removed $($rfwl.Sender) from Sender Whitelist`r"
        }
    }
    Catch{
        Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text "$_.Message indicates an error has occurred.  Please review message and retry once resolved." -Color "Red"
    }
    Clear-Variable rfwl -ErrorAction SilentlyContinue
    Clear-Variable rfwls -ErrorAction SilentlyContinue
})

$SpamFilterRemoveDomainWhitelistButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline
    }
    Clear-Variable rdfwl -ErrorAction SilentlyContinue
    Clear-Variable rdfwls -ErrorAction SilentlyContinue
    Try{
        $rdfwls = Get-HostedContentFilterPolicy Default
        $rdfwls = $rdfwls.AllowedSenderDomains | Select-Object -Property Domain | Out-GridView -Passthru -Title "Select Multiple Senders By Holding Ctrl"
        foreach($rdwfl in $rdfwls){
            Set-HostedContentFilterPolicy Default -AllowedSenderDomains @{Remove="$($rdwfl.Domain)"}
            Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text "Removed $($rdwfl.Domain) from Domain Whitelist`r"
        }
    }
    Catch{
        Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text "$_.Message indicates an error has occurred.  Please review message and retry once resolved."
    }
    Clear-Variable rdfwl -ErrorAction SilentlyContinue
    Clear-Variable rdfwls -ErrorAction SilentlyContinue
})


$SpamFilterRemoveUserBlacklistButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline
    }
    Clear-Variable rfbl -ErrorAction SilentlyContinue
    Clear-Variable rfbls -ErrorAction SilentlyContinue
    Try{
        $rfbls = Get-HostedContentFilterPolicy Default
        $rfbls = $rfbls.BlockedSenders | Select-Object -Property Sender | Out-GridView -Passthru -Title "Select Multiple Senders By Holding Ctrl"
        foreach($rfbl in $rfbls){
            Set-HostedContentFilterPolicy Default -BlockedSenders @{Remove="$($rfbl.sender)"}
            Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text  "Removed $($rfbl.Sender) from Sender Blacklist`r"
        }
    }
    Catch{
        Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text  "$_.Message indicates an error has occurred.  Please review message and retry once resolved." -Color "Red"
    }
    Clear-Variable rfbl -ErrorAction SilentlyContinue
    Clear-Variable rfbls -ErrorAction SilentlyContinue
})

$SpamFilterRemoveDomainBlacklistButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }
    Catch {
        Connect-ExchangeOnline
    }
    Clear-Variable rdfbl -ErrorAction SilentlyContinue
    Clear-Variable rdfbls -ErrorAction SilentlyContinue
    Try{
        $rdfbls = Get-HostedContentFilterPolicy Default
        $rdfbls = $rdfbls.BlockedSenderDomains | Select-Object -Property Domain | Out-GridView -Passthru -Title "Select Multiple Senders By Holding Ctrl"
        foreach($rdfbl in $rdfbls){
            Set-HostedContentFilterPolicy Default -BlockedSenderDomains @{Remove="$($rdfbl.Domain)"}
            Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text "Removed $($rdfbl.Domain) from Domain Blacklist`r"
        }
    }
    Catch{
        Write-RichtextBox -TextBox $SpamFilterRichTextBox -Text "$_.Message indicates an error has occurred.  Please review message and retry once resolved."
    }
    Clear-Variable rdfbl -ErrorAction SilentlyContinue
    Clear-Variable rdfbls -ErrorAction SilentlyContinue
})

### End Spam Filter Tab Functionality

$UserForm.Add_Loaded({
    Try{
        Get-AzureADDomain -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
    }
    Set-Comboboxes
})

$null = $UserForm.ShowDialog()