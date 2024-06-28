from faker import Faker


def generate_fake_user():
    fake = Faker('ru_RU')

    return {
        'name': fake.name(),
        'address': fake.address(),
        'email': fake.email(),
        'phone_number': fake.phone_number(),
        'birth_date': fake.date_of_birth().strftime("%d.%m.%Y"),
        'company': fake.company(),
        'job': fake.job()
    }


def get_fake_users(count: int):
    return [generate_fake_user() for _ in range(count)]
