require 'test_helper'

class ExcelReportsControllerTest < ActionDispatch::IntegrationTest
  setup do
    @excel_report = excel_reports(:one)
  end

  test "should get index" do
    get excel_reports_url
    assert_response :success
  end

  test "should get new" do
    get new_excel_report_url
    assert_response :success
  end

  test "should create excel_report" do
    assert_difference('ExcelReport.count') do
      post excel_reports_url, params: { excel_report: {  } }
    end

    assert_redirected_to excel_report_url(ExcelReport.last)
  end

  test "should show excel_report" do
    get excel_report_url(@excel_report)
    assert_response :success
  end

  test "should get edit" do
    get edit_excel_report_url(@excel_report)
    assert_response :success
  end

  test "should update excel_report" do
    patch excel_report_url(@excel_report), params: { excel_report: {  } }
    assert_redirected_to excel_report_url(@excel_report)
  end

  test "should destroy excel_report" do
    assert_difference('ExcelReport.count', -1) do
      delete excel_report_url(@excel_report)
    end

    assert_redirected_to excel_reports_url
  end
end
