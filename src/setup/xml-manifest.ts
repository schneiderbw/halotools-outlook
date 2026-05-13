// XML manifest generator for the legacy OfficeApp / MailApp schema.
// This is the format Microsoft 365 admin → Integrated apps → "Office Add-in"
// → "Upload manifest file (.xml) from device" actually accepts today, and the
// path that propagates fast and supports admin push-install for users/groups.
//
// We bake the MSP's Halo URL + Client ID into the SourceLocation URL so the
// SPA self-configures on first launch (same trick as the JSON manifest path).

export interface TenantInput {
  haloBaseUrl: string;
  clientId: string;
}

function xmlEscape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function tenantTaskpaneUrl(haloBaseUrl: string, clientId: string): string {
  const origin = window.location.origin;
  const params = new URLSearchParams({ halo: haloBaseUrl, clientId }).toString();
  return `${origin}/outlook/?${params}`;
}

export function buildXmlManifest(input: TenantInput): string {
  const id = crypto.randomUUID();
  const origin = window.location.origin;
  const taskpaneUrl = xmlEscape(tenantTaskpaneUrl(input.haloBaseUrl, input.clientId));
  const haloOrigin = xmlEscape(new URL(input.haloBaseUrl).origin);

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
    xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
    xsi:type="MailApp">
  <Id>${id}</Id>
  <Version>0.1.0</Version>
  <ProviderName>Rising Tide Consulting Group</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Halo for Outlook" />
  <Description DefaultValue="Surface HaloPSA context and log emails to tickets without leaving Outlook." />
  <IconUrl DefaultValue="${origin}/outlook/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${origin}/outlook/assets/icon-128.png" />
  <SupportUrl DefaultValue="https://iusehalo.com" />
  <AppDomains>
    <AppDomain>${haloOrigin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.10">
      <Set Name="Mailbox" MinVersion="1.10" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${taskpaneUrl}" />
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" MinVersion="1.10" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="HaloGroup">
                <Label resid="GroupLabel" />
                <Icon>
                  <bt:Image size="16" resid="Icon16" />
                  <bt:Image size="32" resid="Icon32" />
                  <bt:Image size="80" resid="Icon80" />
                </Icon>
                <Control xsi:type="Button" id="ShowHaloPane">
                  <Label resid="ButtonLabel" />
                  <Supertip>
                    <Title resid="ButtonTitle" />
                    <Description resid="ButtonDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon16" />
                    <bt:Image size="32" resid="Icon32" />
                    <bt:Image size="80" resid="Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon16" DefaultValue="${origin}/outlook/assets/icon-16.png" />
        <bt:Image id="Icon32" DefaultValue="${origin}/outlook/assets/icon-32.png" />
        <bt:Image id="Icon80" DefaultValue="${origin}/outlook/assets/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="${taskpaneUrl}" />
        <bt:Url id="Taskpane.Url" DefaultValue="${taskpaneUrl}" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="HaloPSA" />
        <bt:String id="ButtonLabel" DefaultValue="Halo" />
        <bt:String id="ButtonTitle" DefaultValue="Open Halo" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="ButtonDescription" DefaultValue="Show sender context, open tickets, and log this email to a HaloPSA ticket." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

export function downloadXmlManifest(xml: string, filename: string): void {
  const blob = new Blob([xml], { type: "application/xml" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
