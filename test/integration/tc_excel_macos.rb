# frozen_string_literal: true

require 'tc_helper'
require_relative 'tc_excel_shared'

class TestExcelMacOS < Minitest::Test
  include TestExcelShared

  OSASCRIPT_TIMEOUT_SECONDS = 10

  private

  def setup_excel_application
    excel_app_available?
  end

  def teardown_excel_application
    return unless @excel_available

    # Force quit Excel to ensure clean state
    osascript(<<~APPLESCRIPT)
      try
        tell application "Microsoft Excel"
          quit saving no
        end tell
      on error
        -- Excel might already be closed
      end try
    APPLESCRIPT

    # Give Excel time to fully quit
    sleep 1
  rescue StandardError
  end

  def assert_excel_file_opens(file_path)
    with_workbook(file_path) {} # File opened successfully
    true
  end

  def assert_excel_cell_values(file_path, expected_values)
    with_workbook(file_path) do |workbook_name|
      expected_values.each do |cell_address, expected_value|
        actual_value = get_cell_value(workbook_name, 1, cell_address)

        # Handle type conversions for AppleScript
        actual_value = normalize_cell_value(actual_value)
        expected_value = normalize_cell_value(expected_value)

        assert_equal expected_value, actual_value,
                     "Cell #{cell_address} should contain '#{expected_value}' but contains '#{actual_value}'"
      end
    end
    true
  end

  def assert_excel_cell_values_by_sheet(file_path, sheet_name, expected_values)
    with_workbook(file_path) do |workbook_name|
      expected_values.each do |cell_address, expected_value|
        actual_value = get_cell_value_by_sheet(workbook_name, sheet_name, cell_address)

        # Handle type conversions for AppleScript
        actual_value = normalize_cell_value(actual_value)
        expected_value = normalize_cell_value(expected_value)

        assert_equal expected_value, actual_value,
                     "Cell #{cell_address} in sheet '#{sheet_name}' should contain '#{expected_value}' but contains '#{actual_value}'"
      end
    end
    true
  end

  def with_workbook(file_path)
    absolute_path = File.absolute_path(file_path)
    workbook_name = nil

    # DEBUG: Log what we're trying to open
    puts "DEBUG: Opening file: #{absolute_path}" if ENV['DEBUG']

    # First, ensure Excel is in a good state
    ensure_excel_ready

    # Open the workbook with better error handling
    result = osascript(<<~APPLESCRIPT)
      try
        -- Don't use activate, just work with Excel directly
        tell application "Microsoft Excel"
          -- Disable alerts to prevent dialog boxes
          set display alerts to false

          -- Try to open the workbook
          set workbook_ref to (open workbook workbook file name "#{absolute_path}" with read only)

          -- Get the name for reference
          set workbook_name to name of workbook_ref

          -- Re-enable alerts
          set display alerts to true

          return workbook_name
        end tell
      on error errMsg number errNum
        return "Error " & errNum & ": " & errMsg
      end try
    APPLESCRIPT

    # DEBUG: Log the result
    puts "DEBUG: Open result: #{result}" if ENV['DEBUG']

    if result.include?('Error')
      # Try alternative open method if first fails
      result = alternative_open_workbook(absolute_path)

      if result.include?('Error')
        flunk("Failed to open Excel file '#{file_path}': #{result}")
      end
    end

    workbook_name = result.strip
    yield workbook_name
  ensure
    # Close the workbook
    if workbook_name && !workbook_name.include?('Error')
      osascript(<<~APPLESCRIPT)
        try
          tell application "Microsoft Excel"
            set display alerts to false
            close workbook "#{workbook_name}" saving no
            set display alerts to true
          end tell
        on error
          -- Workbook might already be closed
        end try
      APPLESCRIPT
    end
  end

  def alternative_open_workbook(absolute_path)
    # Try using Finder to open the file
    osascript(<<~APPLESCRIPT)
      try
        -- Use Finder to open the file with Excel
        tell application "Finder"
          open POSIX file "#{absolute_path}" using application file id "com.microsoft.Excel"
        end tell

        delay 2

        -- Now get the workbook name from Excel
        tell application "Microsoft Excel"
          set workbook_name to name of active workbook
          return workbook_name
        end tell
      on error errMsg number errNum
        return "Error " & errNum & ": " & errMsg
      end try
    APPLESCRIPT
  end

  def ensure_excel_ready
    # Make sure Excel is running and ready
    osascript(<<~APPLESCRIPT)
      try
        tell application "Microsoft Excel"
          -- Just check if Excel is responding
          set app_version to version

          -- Close any dialog boxes that might be open
          set display alerts to false

          -- Close all workbooks to start fresh
          close every workbook saving no

          set display alerts to true
        end tell
      on error
        -- Excel might not be running, that's OK
      end try
    APPLESCRIPT
  end

  def get_cell_value(workbook_name, sheet_index, cell_address)
    result = osascript(<<~APPLESCRIPT)
      try
        tell application "Microsoft Excel"
          tell workbook "#{workbook_name}"
            tell worksheet #{sheet_index}
              set cell_value to value of range "#{cell_address}"
              if cell_value is missing value then
                return ""
              else
                return cell_value as string
              end if
            end tell
          end tell
        end tell
      on error errMsg number errNum
        return "Error " & errNum & ": " & errMsg
      end try
    APPLESCRIPT

    if result.include?('Error')
      flunk("Failed to read cell #{cell_address}: #{result}")
    end

    result.strip
  end

  def get_cell_value_by_sheet(workbook_name, sheet_name, cell_address)
    result = osascript(<<~APPLESCRIPT)
      try
        tell application "Microsoft Excel"
          tell workbook "#{workbook_name}"
            tell worksheet "#{sheet_name}"
              set cell_value to value of range "#{cell_address}"
              if cell_value is missing value then
                return ""
              else
                return cell_value as string
              end if
            end tell
          end tell
        end tell
      on error errMsg number errNum
        return "Error " & errNum & ": " & errMsg
      end try
    APPLESCRIPT

    if result.include?('Error')
      flunk("Failed to read cell #{cell_address} from sheet '#{sheet_name}': #{result}")
    end

    result.strip
  end

  def normalize_cell_value(value)
    return nil if value.nil? || value.empty?

    case value
    when /^-?\d+$/
      value.to_i
    when /^-?\d+\.\d+$/
      value.to_f
    when /^missing value$/
      nil
    else
      value.to_s
    end
  end

  def osascript(script)
    # Use macOS timeout command - simple and reliable for macOS-only code
    result = `timeout #{OSASCRIPT_TIMEOUT_SECONDS} osascript -e '#{script.gsub("'", "\\'")}' 2>&1`.strip

    if $CHILD_STATUS.exitstatus == 124 # timeout command exit code
      "Error: Script timed out after #{OSASCRIPT_TIMEOUT_SECONDS} seconds"
    else
      result
    end
  end

  def require_or_skip!
    flunk('This test requires Excel macOS integration') if ENV['CI_EXCEL_MACOS'] && !excel_macos?
    skip('Excel macOS integration tests only run on macOS') unless macos?
    skip('Excel macOS integration tests require Microsoft Excel to be installed') unless excel_macos?
  end

  def excel_macos?
    macos? && excel_app_available?
  end

  def excel_app_available?
    return @excel_app_available if defined?(@excel_app_available)

    # Simpler check without activation
    result = osascript(<<~APPLESCRIPT)
      try
        tell application "Microsoft Excel"
          activate
          delay 2
          quit
        end tell
        return "OK"
      on error errMsg
        return "Error: " & errMsg
      end try
    APPLESCRIPT

    @excel_app_available = result == 'OK'
  rescue StandardError
    @excel_app_available = false
  end
end
