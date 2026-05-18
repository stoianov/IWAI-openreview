import pandas as pd

def merge_email_lists(file1, file2, output_file):

    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # keep only the relevant columns
    df1 = df1[["Name", "Email"]]
    df2 = df2[["Name", "Email"]]

    # concatenate
    df = pd.concat([df1, df2], ignore_index=True)

    # clean emails
    df["Email"] = (df["Email"].astype(str).str.strip().str.lower())

    # remove empty emails
    df = df[df["Email"] != ""]

    # remove duplicates by email
    df = df.drop_duplicates(subset="Email", keep="first")

    # sort alphabetically
    df = df.sort_values(by=["Name", "Email"])

    # export
    df.to_excel(output_file, index=False, engine="openpyxl")

    print(f"Merged list exported to: {output_file}")
    print(f"Total unique emails: {len(df)}")

if __name__ == '__main__':
    merge_email_lists("2025/IWAI2025-Participants.xlsx", "2024/IWAI2024-Participants.xlsx","IWAI-Participants.xlsx")