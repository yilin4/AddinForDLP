1. Clone the repo to local
```git
git clone https://github.com/yilin4/AddinForDLP.git
```
2. Go to the repo root folder and install dependencies.
```git
npm install
```
3. Build the project.
```git
npm run build
```
4. Start the addin.
```git
npm start
```
5. Deploy the add-in with the manifest in this project via [Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps?view=o365-worldwide), or sideload your add-in following the guidance for [Word online](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

6. Create a new Word document or open an existing one, and you will see the headers are added to the document.

> [!NOTE]
> The feature is in preview for early testing. You can not deploy a add-in with this feature to your customer yet. Please also notice the preview version of feature may be different from the released version.
> 
> Supported clients and channels: Office Win32 Desktop DevMain channel insider ring, version>= 16.0.18324.20032 and Offilce online. Office MAC Desktop is not supported yet.
> 
> For more information on this feature, check [this documentation](https://github.com/OfficeDev/office-js-docs-pr/blob/WXP-event-based-activation/docs/develop/WXP-event-based-activation.md)
