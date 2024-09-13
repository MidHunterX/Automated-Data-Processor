import sqlite3  # SQLite DB operations
from function import var


def main():
    db_file = var["db_file"]

    print("ℹ️ Connecting to Database")
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    delete_graduates_sql = """
    DELETE FROM Students
    WHERE Class IN (10, 12, 15, 17);
    """

    verify_graduates_removed_sql = """
    SELECT COUNT(*) FROM Students
    WHERE Class IN (10, 12, 15, 17);
    """

    promote_students_sql = """
    UPDATE Students
    SET Class = Class + 1
    WHERE Class NOT IN (10, 12, 15, 17);
    """

    try:
        cursor.execute(delete_graduates_sql)

        # Verify graduates removal before promoting
        cursor.execute(verify_graduates_removed_sql)
        graduates_remaining = cursor.fetchone()[0]

        if graduates_remaining == 0:
            print("All graduates have been successfully removed.")
        else:
            print(f"There are still {graduates_remaining} graduates remaining")
            return

        cursor.execute(promote_students_sql)
        conn.commit()
        print("Students promoted to next year successfully.")

    except KeyboardInterrupt:
        print("Caught the Keyboard Interrupt ;D")

    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        conn.rollback()

    finally:
        print("ℹ️ Closing DB")
        cursor.close()
        conn.close()


if __name__ == "__main__":
    main()
