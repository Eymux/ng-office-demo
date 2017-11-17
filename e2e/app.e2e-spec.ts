import { NgOfficeDemoPage } from './app.po';

describe('ng-office-demo App', () => {
  let page: NgOfficeDemoPage;

  beforeEach(() => {
    page = new NgOfficeDemoPage();
  });

  it('should display welcome message', () => {
    page.navigateTo();
    expect(page.getParagraphText()).toEqual('Welcome to app!!');
  });
});
