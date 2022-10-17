#include<iostream>
#include<stack>
#include<libxl.h>

#define INF 99999
#define station 121

using namespace std;
using namespace libxl;



typedef struct edge {
	int id;
	int time;
	int tran;
	edge* next = NULL;
};

typedef struct Node {						// Node
	string* name;							// 역 이름
	double dis;								// 거리
	edge* graph;
	int n;
	int* root;
	Node* next = NULL;
};

Node* ALLsub;	//ALL


int choose(bool t)
{
	int min = INF;
	int min_t = INF;
	int pos = -1;
	if (t)
	{
		for (int i = 0; i < ALLsub->n; i++)
		{
			if (ALLsub->root[i] < 0)
			{
				if (ALLsub->graph[i].time < min)
				{
					min = ALLsub->graph[i].tran;
					pos = i;
				}
				else if (ALLsub->graph[i].tran == min)
				{
					if (ALLsub->graph[i].tran < min_t)
					{
						min = ALLsub->graph[i].time;
						min_t = ALLsub->graph[i].tran;
						pos = i;
					}
				}
			}
		}
	}
	return pos;
}

void Dijk(int start, bool t)
{
	int minpos;
	for (int i = 0; i < ALLsub->n; i++)
	{
		ALLsub->root[i] = -1;
		ALLsub->graph[i].id = -1;
		ALLsub->graph[i].time = INF;
		ALLsub->graph[i].tran = INF;
	}
	for (edge* p = ALLsub->graph[start].next; p != NULL; p = p->next)
	{
		if (ALLsub->name[start] == ALLsub->name[p->id])
		{
			ALLsub->graph[p->id].time = 0;
			ALLsub->graph[p->id].tran = 0;
		}
		else
		{
			ALLsub->graph[p->id].time = p->time;
			ALLsub->graph[p->id].tran = p->tran;
		}
	}

	for (int i = 0; i < ALLsub->n; i++)
	{
		if (ALLsub->name[start] == ALLsub->name[i])
		{
			for (edge* p = ALLsub->graph[i].next; p != NULL; p = p->next)
			{
				if (ALLsub->name[i] == ALLsub->name[p->id])
				{
					p->time = 0;
					p->tran = 0;
				}
			}
		}
	}
	ALLsub->root[start] = 0;
	ALLsub->graph[start].time = 0;
	ALLsub->graph[start].tran = 0;

	for (int i = 0, prev = start; i < ALLsub->n - 2; i++)
	{
		minpos = choose(t);
		ALLsub->root[minpos] = 0;
		for (edge* p = ALLsub->graph[minpos].next; p != NULL; p = p->next)
		{
			if (t)
			{
				if (ALLsub->graph[minpos].time < ALLsub->graph[p->id].time)
				{
					ALLsub->graph[p->id].time = ALLsub->graph[minpos].time + p->time;
					ALLsub->graph[p->id].tran = ALLsub->graph[minpos].tran + p->tran;
				}
				else if (ALLsub->graph[minpos].tran + p->tran < ALLsub->graph[p->id].tran)
				{
					ALLsub->graph[p->id].time = ALLsub->graph[minpos].time + p->time;
					ALLsub->graph[p->id].tran = ALLsub->graph[minpos].tran + p->tran;
				}
			}
		}
	}
}



void* InitFile(int N)		// 지하철 호선, 행
{
	edge* ta;
	edge* p;
	ALLsub->n = N;

	ALLsub->graph = new edge[ALLsub->n];
	ALLsub->name = new string[ALLsub->n];
	ALLsub->root = new int[ALLsub->n];

	int t1, time;
	string sta;
	int coi = 1;
	int num = 0;
	Book* book = xlCreateBook();
	if (book)
	{
		if (book->load("ALLsub.xls"))
		{
			Sheet* sheet = book->getSheet(0);
			if (sheet)
			{
				while (1)
				{
					const char* s = sheet->readStr(num, 0);
					double d = sheet->readNum(num, 1);
					double q = sheet->readNum(num, 2);

					if (s == NULL && d == NULL) {

						break;
					}
					else if (s == NULL)
					{
						s = "NoName";
					}
					else if (d == NULL)
					{
						d = 0;
					}
					t1 = coi;
					sta = s;
					time = q;

					t1 -= 1;
					ta = new edge();
					ta->id = t1;
					ta->time = time;
					coi++;
					num++;

					p = &ALLsub->graph[t1];
					while (p->next != NULL)
					{
						p = p->next;
					}
					p->next = ta;
					ALLsub->name[t1] = sta;
				}
			}
		}
		book->release();
	}
	return ALLsub;
}

bool Fineroot(stack<int>& a, int start, int end)
{
	a.push(start);
	if (start == end)
	{
		return true;
	}

	for (edge* p = ALLsub->graph[start].next; p != NULL; p = p->next)
	{
		if (ALLsub->graph[start].time + p->time == ALLsub->graph[p->id].time && ALLsub->graph[start].tran + p->tran == ALLsub->graph[p->id].tran)
		{
			if (!Fineroot(a, p->id, end))
			{
				a.pop();
			}
			else
			{
				return true;
			}
		}
	}
	return false;
}
void print()
{

	int start = -1, end = -1;
	int min, v, prev;
	string sy;
	stack<int>a, b;
	while (start == end)
	{
		while (start < 0)
		{
			cout << "출발역을 입력하세요 :";
			cin >> sy;
			cout << endl;

			for (int i = 0; ALLsub->n; i++)
			{
				if (ALLsub->name[i] == sy)
				{
					start = i;
					break;
				}
			}

			if (start < 0)
			{
				cout << "잘못입력했습니다" << endl;
			}
		}
		while (end < 0)
		{
			cout << "도착역 입력 :";
			cin >> sy;
			cout << endl;

			for (int i = 0; ALLsub->n; i++)
			{
				if (ALLsub->name[i] == sy)
				{
					end = i;
					break;
				}
			}
			if (end < 0)
			{
				cout << "잘못입력했습니다" << endl;
			}
		}
		if (start == end)
		{
			cout << "출발역=도착역" << endl;
			start = end = -1;
		}
	}
	cout << "경로 :" << ALLsub->name[start] << " -> " << ALLsub->name[end] << "\n" << endl;
	Dijk(start, true);

	min = ALLsub->graph[end].time;
	for (int i = 0; i < ALLsub->n; i++)
	{
		if (ALLsub->name[end] == ALLsub->name[i])
		{
			if (ALLsub->graph[i].time < min)
			{
				min = ALLsub->graph[i].time;
				end = i;
			}
		}
	}
	Fineroot(a, start, end);

	while (!a.empty())
	{
		b.push(a.top());
		a.pop();
	}
	for (v = 0, prev = -1; !b.empty(); v++)
	{
		if (prev != -1 && ALLsub->name[prev] == ALLsub->name[b.top()])
		{
			if (ALLsub->name[b.top()] != ALLsub->name[start])
				cout << "<환승>";

			v--;
			prev = b.top();
			b.pop();
		}
		else
		{
			if (v != 0)
				cout << " -> ";
			else
				cout << " ";

			cout << ALLsub->name[b.top()];
			prev = b.top();
			b.top();
		}
	}
}

typedef struct ipa {						// ipa xl 테스트
	string name1;
	double dis1;
	double edge1;
	ipa* next1 = NULL;
};

ipa* ta;

ipa* Init(ipa* subway)		// 지하철 호선, 행 xl 테스트
{
	int num = 0;

	Book* book = xlCreateBook();
	if (book)
	{
		if (book->load("Subway.xls"))
		{
			Sheet* sheet = book->getSheet(0);
			if (sheet)
			{
				while (1)
				{
					const char* s = sheet->readStr(num, 0);
					double d = sheet->readNum(num, 1);
					double r = sheet->readNum(num, 2);
					if (s == NULL && d == NULL) {

						break;
					}
					else if (s == NULL)
					{
						s = "NoName";
					}
					else if (d == NULL)
					{
						d = 0;
					}
					ipa* temp = new ipa;
					temp->name1 = s;
					temp->dis1 = d;
					temp->edge1 = r;
					temp->next1 = NULL;

					if (subway == NULL)
					{
						subway = temp;
					}
					else
					{
						ipa* nod = subway;
						while (nod->next1 != NULL)
						{
							nod = nod->next1;
						}
						nod->next1 = temp;
					}
					num++;
				}
			}
		}
		book->release();
	}
	return subway;
}

void prinf(ipa* subway)
{
	while (subway != NULL)
	{
		cout.width(15);
		cout << subway->name1;
		cout.width(10);
		cout << subway->dis1;
		cout.width(10);
		cout << subway->edge1 << endl;
		subway = subway->next1;
	}
}


int main()
{
	ta = Init(ta);
	prinf(ta);
	cout << endl;
	print();
}