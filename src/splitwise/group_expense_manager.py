from typing import Dict
from splitwise import Splitwise, Expense
from splitwise.user import ExpenseUser


class GroupExpenseManager:
    """
    Manager to adding expenses into given group
    """

    def __init__(
            self,
            splitwise_connection: Splitwise,
            group_id: int,
    ):
        """

        :param splitwise_connection: authorized splitwise connection
        :param group_id: group id to add expenses
        """
        self._splitwise_connection = splitwise_connection
        self.group_id = group_id

    def add_expense(
            self,
            description: str,
            users_shares: Dict[int, float],
            master_user_id,
    ):
        """

        :param description: text to write on splitwise
        :param users_shares: mapped users id on their owed shares
        :param master_user_id: Person who paid for everybody
        :return:
        """
        total_cost = sum(users_shares.values())

        expense = Expense()
        expense.setCost(total_cost)
        expense.setDescription(description)

        users = []

        for user_id in users_shares.keys():
            user = ExpenseUser()
            user.setId(user_id)
            user.setOwedShare(users_shares[user_id])

            if user_id == master_user_id:
                user.setPaidShare(total_cost)
            users.append(user)

        if master_user_id not in users_shares.keys():
            user = ExpenseUser()
            user.setId(master_user_id)
            user.setPaidShare(total_cost)

        expense.setUsers(users)
        expense.setCurrencyCode('RUB')
        expense.setGroupId(self.group_id)
        self._splitwise_connection.createExpense(expense)
