from generator_core import get_schedule_suffixes_for_country

# Test the schedule settings functionality
def test_schedule_settings():
    # Test default behavior
    default_schedules = ["Mon-Fri 9-17", "Mon-Sun 24/7"]
    
    # Empty per-country settings should return default
    result1 = get_schedule_suffixes_for_country("DE", "HS DE", {}, default_schedules)
    print(f"Test 1 - DE with no per-country settings: {result1}")
    assert result1 == default_schedules
    
    # Test with per-country settings
    per_country_settings = {
        "HS PL": "Mon-Fri 8-16\nSat 10-14",
        "DS PL": "Mon-Fri 9-17\nSun 12-18",
        "DE": "Mon-Thu 8-17\nFri 8-15"
    }
    
    # Test HS PL
    result2 = get_schedule_suffixes_for_country("PL", "HS PL", per_country_settings, default_schedules)
    print(f"Test 2 - HS PL with custom settings: {result2}")
    assert result2 == ["Mon-Fri 8-16", "Sat 10-14"]
    
    # Test DS PL
    result3 = get_schedule_suffixes_for_country("PL", "DS PL", per_country_settings, default_schedules)
    print(f"Test 3 - DS PL with custom settings: {result3}")
    assert result3 == ["Mon-Fri 9-17", "Sun 12-18"]
    
    # Test DE (no receiver needed for DE)
    result4 = get_schedule_suffixes_for_country("DE", "HS DE", per_country_settings, default_schedules)
    print(f"Test 4 - DE with custom settings: {result4}")
    assert result4 == ["Mon-Thu 8-17", "Fri 8-15"]
    
    # Test country not in settings (should return default)
    result5 = get_schedule_suffixes_for_country("UA", "DS UA", per_country_settings, default_schedules)
    print(f"Test 5 - UA not in settings: {result5}")
    assert result5 == default_schedules
    
    print("All tests passed!")

if __name__ == "__main__":
    test_schedule_settings()
