import pandas as pd
import sys
from shopify_feed_generator import generate_shopify_feed

print("Testing finish prioritization in v1.5.0")

# Test case 1: Cadiz product (should use product-specific finishes)
print("\nTest Case 1: Cadiz product (rows 14786-14787)")
output_df1 = generate_shopify_feed('MASTER COPY.xlsx', 'test_cadiz.xlsx', test_mode=True)
if 'CONFIG' in sys.modules['shopify_feed_generator'].__dict__:
    sys.modules['shopify_feed_generator'].CONFIG["test_start_row"] = 14786
    sys.modules['shopify_feed_generator'].CONFIG["test_end_row"] = 14787

# Test case 2: Product with finish count but no specific name
# You'll need to replace these row numbers with appropriate ones
print("\nTest Case 2: Product with finish count (You may need to replace these rows)")
if 'CONFIG' in sys.modules['shopify_feed_generator'].__dict__:
    sys.modules['shopify_feed_generator'].CONFIG["test_start_row"] = 14770
    sys.modules['shopify_feed_generator'].CONFIG["test_end_row"] = 14774
output_df2 = generate_shopify_feed('MASTER COPY.xlsx', 'test_finish_count.xlsx', test_mode=True)

# Analyze the results
print("\nAnalyzing results:")
print(f"Test 1 (Cadiz): Generated {len(output_df1)} rows")
unique_finishes1 = output_df1['Option2 Value'].unique()
print(f"Found {len(unique_finishes1)} unique finishes in Test 1:")
for i, finish in enumerate(unique_finishes1):
    print(f"  {i+1}. {finish}")

print(f"\nTest 2 (Finish Count): Generated {len(output_df2)} rows")
unique_finishes2 = output_df2['Option2 Value'].unique()
print(f"Found {len(unique_finishes2)} unique finishes in Test 2:")
for i, finish in enumerate(unique_finishes2):
    print(f"  {i+1}. {finish}")

print("\nVerification complete!") 